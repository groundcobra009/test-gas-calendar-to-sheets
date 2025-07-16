
/**
 * カレンダーバックアップシステム
 * コンテナバインドスプレッドシート使用版
 * 安全性を重視し、カレンダーID指定必須
 */

// 設定
const CONFIG = {
    BACKUP_MONTHS_FUTURE: 12, // 未来何ヶ月分をバックアップするか
    BACKUP_TIME_HOUR: 5,      // バックアップ実行時刻（時）
    SHEET_PREFIX: 'Calendar_',
    MAX_EVENTS_PER_BATCH: 100, // バッチ処理での最大イベント数
    EXECUTION_TIME_LIMIT: 5 * 60 * 1000 // 実行時間制限（5分）
};

/**
 * メニューを作成
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📅 カレンダーバックアップ')
        .addItem('🔧 初期設定（全期間バックアップ）', 'initialFullBackup')
        .addItem('🔄 今日以降のバックアップ', 'dailyBackup')
        .addSeparator()
        .addItem('⏰ 自動バックアップ設定', 'setupTrigger')
        .addItem('🛑 自動バックアップ停止', 'deleteTrigger')
        .addSeparator()
        .addItem('📋 カレンダー一覧表示', 'showCalendarList')
        .addItem('⚙️ 設定', 'showSettings')
        .addItem('🗑️ 設定リセット', 'resetSettings')
        .addToUi();
}

/**
 * 初回設定 - 全期間バックアップ（コンテナバインド版）
 */
function initialFullBackup() {
    const ui = SpreadsheetApp.getUi();
    const startTime = new Date().getTime();

    try {
        // カレンダー選択
        const calendars = getCalendarList();
        if (calendars.length === 0) {
            ui.alert('エラー', 'アクセス可能なカレンダーが見つかりません。', ui.ButtonSet.OK);
            return;
        }

        // カレンダー選択ダイアログ
        const calendarId = selectCalendar(calendars);
        if (!calendarId) return;

        // 実行前確認
        const calendar = CalendarApp.getCalendarById(calendarId);
        if (!calendar) {
            ui.alert('エラー', '選択されたカレンダーにアクセスできません。', ui.ButtonSet.OK);
            return;
        }

        // 期間設定
        const startDate = new Date(2012, 0, 1); // 2012年1月1日から
        const endDate = new Date();
        endDate.setFullYear(endDate.getFullYear() + 1); // 1年後まで

        // 初期設定時の既存データ削除確認
        const clearResponse = ui.alert(
            '初期設定確認',
            '初期設定を実行します。\n\n【重要】既存のバックアップデータの処理方法を選択してください：\n\n「はい」：すべての既存シートを削除して完全に初期化\n「いいえ」：選択したカレンダーのシートのみクリア\n「キャンセル」：処理を中止',
            ui.ButtonSet.YES_NO_CANCEL
        );
        
        if (clearResponse === ui.Button.CANCEL) return;
        
        // 全シート削除が選択された場合
        if (clearResponse === ui.Button.YES) {
            const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
            const sheets = spreadsheet.getSheets();
            
            // 情報シート以外のすべてのシートを削除
            sheets.forEach(sheet => {
                const sheetName = sheet.getName();
                if (!sheetName.includes('バックアップ情報') && sheetName !== 'シート1') {
                    spreadsheet.deleteSheet(sheet);
                }
            });
            
            // 情報シートがある場合はクリア
            const infoSheet = spreadsheet.getSheetByName('バックアップ情報');
            if (infoSheet) {
                const lastRow = infoSheet.getLastRow();
                if (lastRow > 1) {
                    infoSheet.getRange(2, 1, lastRow - 1, 4).clear();
                }
            }
            
            console.log('全シート削除完了');
        }

        // 事前チェック: イベント数推定
        const estimatedEvents = estimateEventCount(calendar, startDate, endDate);
        console.log(`推定イベント数: ${estimatedEvents}件`);
        
        if (estimatedEvents > 500) {
            const response = ui.alert(
                '確認', 
                `大量のイベント（推定${estimatedEvents}件）が検出されました。\n処理に時間がかかる可能性があります。続行しますか？`, 
                ui.ButtonSet.YES_NO
            );
            if (response !== ui.Button.YES) return;
        }

        // バックアップ実行（コンテナバインド版）
        const result = backupCalendarToContainerSheet(calendarId, startDate, endDate, true, startTime);

        if (result.success) {
            ui.alert('完了', `初期バックアップが完了しました。\n${result.eventCount}件のイベントをバックアップしました。\n実行時間: ${result.executionTime}秒\n\nシート: ${result.sheetName}`, ui.ButtonSet.OK);
        } else {
            ui.alert('エラー', `バックアップに失敗しました：${result.error}`, ui.ButtonSet.OK);
        }

    } catch (error) {
        console.error('初期バックアップエラー:', error);
        ui.alert('エラー', `予期しないエラーが発生しました：${error.message}`, ui.ButtonSet.OK);
    }
}

/**
 * イベント数推定
 */
function estimateEventCount(calendar, startDate, endDate) {
    try {
        // 1週間分のサンプルを取得して推定
        const sampleEndDate = new Date(startDate.getTime() + 7 * 24 * 60 * 60 * 1000);
        const sampleEvents = calendar.getEvents(startDate, sampleEndDate);
        const totalDays = Math.ceil((endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000));
        return Math.floor((sampleEvents.length / 7) * totalDays);
    } catch (error) {
        console.error('イベント数推定エラー:', error);
        return 0;
    }
}

/**
 * コンテナバインドスプレッドシートでのカレンダーバックアップ
 */
function backupCalendarToContainerSheet(calendarId, startDate, endDate, isFullBackup, startTime) {
    try {
        // カレンダー取得
        const calendar = CalendarApp.getCalendarById(calendarId);
        if (!calendar) {
            return { success: false, error: 'カレンダーが見つかりません' };
        }

        // コンテナバインドスプレッドシートを取得
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        if (!spreadsheet) {
            return { success: false, error: 'コンテナバインドスプレッドシートが見つかりません' };
        }

        // シート準備
        const sheetResult = getOrCreateSheetInContainer(spreadsheet, calendar);
        if (!sheetResult.success) {
            return { success: false, error: `シート作成エラー: ${sheetResult.error}` };
        }
        const sheet = sheetResult.sheet;

        // 既存データをクリア（全期間バックアップの場合）
        if (isFullBackup) {
            const lastRow = sheet.getLastRow();
            if (lastRow > 1) {
                sheet.getRange(2, 1, lastRow - 1, 8).clear();
            }
        }

        // バッチ処理でイベント取得・保存
        let totalEvents = 0;
        let currentDate = new Date(startDate);
        const batchSize = 30; // 30日ずつ処理

        while (currentDate < endDate) {
            // 時間制限チェック
            if (new Date().getTime() - startTime > CONFIG.EXECUTION_TIME_LIMIT) {
                console.log('実行時間制限により処理を中断');
                break;
            }

            const batchEndDate = new Date(Math.min(
                currentDate.getTime() + batchSize * 24 * 60 * 60 * 1000,
                endDate.getTime()
            ));

            console.log(`処理中: ${currentDate.toLocaleDateString()} - ${batchEndDate.toLocaleDateString()}`);

            try {
                const events = calendar.getEvents(currentDate, batchEndDate);
                
                if (events.length > 0) {
                    const eventData = events.map(event => [
                        event.getTitle(),
                        event.getStartTime(),
                        event.getEndTime(),
                        event.getLocation() || '',
                        event.getDescription() || '',
                        event.getGuestList().map(guest => guest.getEmail()).join(', '),
                        event.getDateCreated(),
                        event.getId()
                    ]);

                    // データを追加
                    const lastRow = sheet.getLastRow();
                    sheet.getRange(lastRow + 1, 1, eventData.length, 8).setValues(eventData);
                    totalEvents += events.length;
                }
            } catch (batchError) {
                console.error(`バッチ処理エラー (${currentDate.toLocaleDateString()}):`, batchError);
            }

            currentDate = batchEndDate;
        }

        // カレンダーIDを保存
        saveCalendarId(calendarId);

        // バックアップ情報を更新
        updateBackupInfoInContainer(spreadsheet, calendar, totalEvents);

        const executionTime = Math.round((new Date().getTime() - startTime) / 1000);
        return { 
            success: true, 
            eventCount: totalEvents, 
            executionTime: executionTime,
            sheetName: sheet.getName()
        };

    } catch (error) {
        console.error('コンテナバックアップエラー:', error);
        return { success: false, error: error.message };
    }
}

/**
 * コンテナ内でのシート取得・作成
 */
function getOrCreateSheetInContainer(spreadsheet, calendar) {
    try {
        // 基本シート名を生成
        let baseSheetName = CONFIG.SHEET_PREFIX + calendar.getName().replace(/[\/\\\?\*\[\]:]/g, '_');
        
        // 長すぎる場合は短縮
        if (baseSheetName.length > 30) {
            baseSheetName = baseSheetName.substring(0, 27) + '...';
        }

        // 既存シートをチェック
        let sheet = spreadsheet.getSheetByName(baseSheetName);
        
        if (!sheet) {
            // シート作成
            sheet = spreadsheet.insertSheet(baseSheetName);
            
            // ヘッダー作成
            const headers = [
                'タイトル', '開始日時', '終了日時', '場所', '説明', 'ゲスト', '作成日', 'ID'
            ];
            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            
            // ヘッダースタイル設定
            const headerRange = sheet.getRange(1, 1, 1, headers.length);
            headerRange.setFontWeight('bold');
            headerRange.setBackground('#4285f4');
            headerRange.setFontColor('#ffffff');
            
            // 列幅調整
            sheet.setColumnWidth(1, 200); // タイトル
            sheet.setColumnWidth(2, 150); // 開始日時
            sheet.setColumnWidth(3, 150); // 終了日時
            sheet.setColumnWidth(4, 120); // 場所
            sheet.setColumnWidth(5, 300); // 説明
            sheet.setColumnWidth(6, 200); // ゲスト
            sheet.setColumnWidth(7, 150); // 作成日
            sheet.setColumnWidth(8, 200); // ID

            console.log(`新規シート作成: ${baseSheetName}`);
        } else {
            console.log(`既存シート使用: ${baseSheetName}`);
        }

        return { success: true, sheet: sheet };

    } catch (error) {
        console.error('コンテナシート作成エラー:', error);
        return { success: false, error: error.message };
    }
}

/**
 * コンテナ内でのバックアップ情報更新
 */
function updateBackupInfoInContainer(spreadsheet, calendar, eventCount) {
    try {
        // 情報シートを取得または作成
        let infoSheet = spreadsheet.getSheetByName('📋 バックアップ情報');
        if (!infoSheet) {
            infoSheet = spreadsheet.insertSheet('📋 バックアップ情報');
            setupInfoSheet(infoSheet);
        }

        // バックアップ情報テーブルの場所を探す
        const data = infoSheet.getDataRange().getValues();
        let targetRow = -1;
        
        for (let i = 0; i < data.length; i++) {
            if (data[i][0] === '📊 バックアップ対象カレンダー') {
                targetRow = i + 2; // ヘッダーの次の行
                break;
            }
        }

        if (targetRow > 0) {
            // 既存エントリをチェック
            let existingRow = -1;
            for (let i = targetRow; i < data.length; i++) {
                if (data[i][1] === calendar.getId()) {
                    existingRow = i + 1; // 1ベースの行番号
                    break;
                }
            }

            const updateData = [
                calendar.getName(),
                calendar.getId(),
                new Date().toLocaleString('ja-JP'),
                `${eventCount}件`
            ];

            if (existingRow > 0) {
                // 既存エントリを更新
                infoSheet.getRange(existingRow, 1, 1, 4).setValues([updateData]);
            } else {
                // 新しいエントリを追加
                const newRow = infoSheet.getLastRow() + 1;
                infoSheet.getRange(newRow, 1, 1, 4).setValues([updateData]);
            }
            
            console.log(`バックアップ情報更新: ${calendar.getName()} (${eventCount}件)`);
        }
    } catch (error) {
        console.error('コンテナバックアップ情報更新エラー:', error);
    }
}

/**
 * 今日以降のイベント更新（削除→再取得）
 */
function updateTodayAndFutureEvents(calendarId) {
    try {
        const calendar = CalendarApp.getCalendarById(calendarId);
        if (!calendar) {
            return { success: false, error: 'カレンダーが見つかりません' };
        }

        // コンテナバインドスプレッドシートを取得
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        if (!spreadsheet) {
            return { success: false, error: 'コンテナバインドスプレッドシートが見つかりません' };
        }

        // 対象シートを取得または作成
        const sheetResult = getOrCreateSheetInContainer(spreadsheet, calendar);
        if (!sheetResult.success) {
            return { success: false, error: `シート準備エラー: ${sheetResult.error}` };
        }
        const sheet = sheetResult.sheet;

        // 今日以降のデータを削除
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        clearTodayAndFutureData(sheet, today);

        // 今日以降のイベントを再取得
        const endDate = new Date();
        endDate.setFullYear(endDate.getFullYear() + 1); // 1年後まで

        const events = calendar.getEvents(today, endDate);
        
        if (events.length > 0) {
            const eventData = events.map(event => [
                event.getTitle(),
                event.getStartTime(),
                event.getEndTime(),
                event.getLocation() || '',
                event.getDescription() || '',
                event.getGuestList().map(guest => guest.getEmail()).join(', '),
                event.getDateCreated(),
                event.getId()
            ]);

            // データを追加
            const lastRow = sheet.getLastRow();
            sheet.getRange(lastRow + 1, 1, eventData.length, 8).setValues(eventData);
        }

        // バックアップ情報を更新
        updateBackupInfoInContainer(spreadsheet, calendar, events.length);

        console.log(`更新完了: ${calendar.getName()} (${events.length}件)`);
        return { 
            success: true, 
            eventCount: events.length, 
            sheetName: sheet.getName() 
        };

    } catch (error) {
        console.error('今日以降のイベント更新エラー:', error);
        return { success: false, error: error.message };
    }
}

/**
 * 今日以降のデータを削除
 */
function clearTodayAndFutureData(sheet, today) {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return; // ヘッダーのみの場合

    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const rowsToDelete = [];

    for (let i = 0; i < data.length; i++) {
        const startDate = new Date(data[i][1]); // 開始日時列
        if (startDate >= today) {
            rowsToDelete.push(i + 2); // 実際の行番号（ヘッダー分+1、0ベース分+1）
        }
    }

    // 後ろから削除（行番号がずれないように）
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        sheet.deleteRow(rowsToDelete[i]);
    }
}

/**
 * 情報シートのセットアップ（コンテナ版・4列対応）
 */
function setupInfoSheet(sheet) {
    // ヘッダー設定
    const headers = [
        ['カレンダーバックアップシステム（コンテナバインド版）'],
        [''],
        ['作成日時', new Date().toLocaleString('ja-JP')],
        ['バックアップ範囲', `今日から${CONFIG.BACKUP_MONTHS_FUTURE}ヶ月後まで`],
        ['自動実行時刻', `毎日${CONFIG.BACKUP_TIME_HOUR}時`],
        ['運用方法', '今日以降のデータを削除→再取得'],
        [''],
        ['📝 使い方'],
        ['1. メニューから「初期設定」を実行'],
        ['2. バックアップしたいカレンダーを選択'],
        ['3. 「自動バックアップ設定」で毎日の自動実行を有効化'],
        ['4. 日次実行で今日以降のデータが更新されます'],
        [''],
        ['⚠️ 注意事項'],
        ['・このスプレッドシートに直接データが保存されます'],
        ['・カレンダーIDを必ず指定してください'],
        ['・プライマリカレンダーの操作は慎重に行ってください'],
        ['・バックアップデータは各カレンダー毎に別シートに保存されます'],
        [''],
        ['📊 バックアップ対象カレンダー'],
        ['カレンダー名', 'カレンダーID', '最終バックアップ日時', 'イベント数']
    ];

    // データを設定
    sheet.getRange(1, 1, headers.length, 4).setValues(headers);

    // スタイル設定
    sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
    sheet.getRange(8, 1).setFontWeight('bold').setBackground('#e8f4fd');
    sheet.getRange(14, 1).setFontWeight('bold').setBackground('#fef7e0');
    sheet.getRange(20, 1).setFontWeight('bold').setBackground('#e8f5e8');
    sheet.getRange(21, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');

    // 列幅調整
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 180);
    sheet.setColumnWidth(4, 100);
}

/**
 * カレンダーバックアップ（既存版、下位互換用）
 */
function backupCalendar(calendarId, startDate, endDate, isFullBackup) {
    try {
        // コンテナバインド版にリダイレクト
        const startTime = new Date().getTime();
        if (isFullBackup) {
            return backupCalendarToContainerSheet(calendarId, startDate, endDate, true, startTime);
        } else {
            return updateTodayAndFutureEvents(calendarId);
        }
    } catch (error) {
        console.error('バックアップエラー:', error);
        return { success: false, error: error.message };
    }
}

/**
 * 日次バックアップ（今日以降のみ）- コンテナバインド版
 */
function dailyBackup() {
    const ui = SpreadsheetApp.getUi();

    try {
        // 設定されたカレンダーIDを取得
        const calendarIds = getStoredCalendarIds();
        if (calendarIds.length === 0) {
            ui.alert('設定エラー', '初期設定を先に実行してください。', ui.ButtonSet.OK);
            return;
        }

        let totalEvents = 0;
        const errors = [];
        const processedSheets = [];

        // 各カレンダーをバックアップ
        for (const calendarId of calendarIds) {
            try {
                const calendar = CalendarApp.getCalendarById(calendarId);
                if (!calendar) {
                    errors.push(`${calendarId}: カレンダーが見つかりません`);
                    continue;
                }

                // 今日以降のイベントを削除→再取得
                const result = updateTodayAndFutureEvents(calendarId);
                if (result.success) {
                    totalEvents += result.eventCount;
                    processedSheets.push(result.sheetName);
                } else {
                    errors.push(`${calendar.getName()}: ${result.error}`);
                }
            } catch (error) {
                errors.push(`${calendarId}: ${error.message}`);
            }
        }

        // 結果表示
        let message = `日次バックアップが完了しました。\n${totalEvents}件のイベントを処理しました。`;
        if (processedSheets.length > 0) {
            message += `\n\n更新されたシート:\n${processedSheets.join('\n')}`;
        }
        if (errors.length > 0) {
            message += `\n\nエラー:\n${errors.join('\n')}`;
        }

        ui.alert('完了', message, ui.ButtonSet.OK);

    } catch (error) {
        console.error('日次バックアップエラー:', error);
        ui.alert('エラー', `予期しないエラーが発生しました：${error.message}`, ui.ButtonSet.OK);
    }
}

/**
 * テスト用：コンテナバインド版の動作確認
 */
function testContainerBackup() {
    console.log('=== コンテナバインド版テスト開始 ===');
    
    try {
        // 1. スプレッドシート確認
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        if (!spreadsheet) {
            console.log('❌ コンテナバインドスプレッドシートが見つかりません');
            return;
        }
        
        console.log(`✅ コンテナスプレッドシート: ${spreadsheet.getName()}`);
        console.log(`   ID: ${spreadsheet.getId()}`);

        // 2. カレンダー一覧取得テスト
        console.log('\n2. カレンダー一覧取得テスト');
        const calendars = getCalendarList();
        console.log(`取得したカレンダー数: ${calendars.length}`);
        calendars.forEach((cal, index) => {
            console.log(`  ${index + 1}. ${cal.name} ${cal.isPrimary ? '[Primary]' : ''}`);
        });

        if (calendars.length === 0) {
            console.log('❌ カレンダーが見つかりません');
            return;
        }

        // 3. プライマリカレンダーでシート作成テスト
        console.log('\n3. シート作成テスト');
        const primaryCalendar = CalendarApp.getDefaultCalendar();
        const sheetResult = getOrCreateSheetInContainer(spreadsheet, primaryCalendar);
        if (!sheetResult.success) {
            console.log(`❌ シート作成失敗: ${sheetResult.error}`);
            return;
        }
        console.log(`✅ シート作成完了: ${sheetResult.sheet.getName()}`);

        // 4. 今日以降の更新テスト
        console.log('\n4. 今日以降の更新テスト');
        const updateResult = updateTodayAndFutureEvents(primaryCalendar.getId());
        if (updateResult.success) {
            console.log(`✅ 更新成功: ${updateResult.eventCount}件のイベント`);
            console.log(`   シート: ${updateResult.sheetName}`);
        } else {
            console.log(`❌ 更新失敗: ${updateResult.error}`);
        }

        console.log('\n=== コンテナバインド版テスト完了 ===');

    } catch (error) {
        console.error('❌ テスト中にエラーが発生:', error);
    }
}

/**
 * カレンダー一覧を取得
 */
function getCalendarList() {
    const calendars = CalendarApp.getAllCalendars();
    return calendars.map(calendar => ({
        id: calendar.getId(),
        name: calendar.getName(),
        isPrimary: calendar.getId() === CalendarApp.getDefaultCalendar().getId()
    }));
}

/**
 * カレンダー選択ダイアログ
 */
function selectCalendar(calendars) {
    const ui = SpreadsheetApp.getUi();

    let message = 'バックアップするカレンダーを選択してください:\n\n';
    calendars.forEach((cal, index) => {
        const primaryMark = cal.isPrimary ? ' [プライマリ]' : '';
        message += `${index + 1}. ${cal.name}${primaryMark}\n`;
    });
    message += '\n番号を入力してください:';

    const response = ui.prompt('カレンダー選択', message, ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() === ui.Button.OK) {
        const index = parseInt(response.getResponseText()) - 1;
        if (index >= 0 && index < calendars.length) {
            return calendars[index].id;
        } else {
            ui.alert('エラー', '無効な番号です。', ui.ButtonSet.OK);
            return null;
        }
    }

    return null;
}

/**
 * カレンダーIDを保存
 */
function saveCalendarId(calendarId) {
    const properties = PropertiesService.getScriptProperties();
    const existingIds = properties.getProperty('CALENDAR_IDS');

    let calendarIds = [];
    if (existingIds) {
        calendarIds = JSON.parse(existingIds);
    }

    if (!calendarIds.includes(calendarId)) {
        calendarIds.push(calendarId);
        properties.setProperty('CALENDAR_IDS', JSON.stringify(calendarIds));
    }
}

/**
 * 保存されたカレンダーIDを取得
 */
function getStoredCalendarIds() {
    const properties = PropertiesService.getScriptProperties();
    const calendarIds = properties.getProperty('CALENDAR_IDS');
    return calendarIds ? JSON.parse(calendarIds) : [];
}

/**
 * 自動バックアップトリガーを設定
 */
function setupTrigger() {
    const ui = SpreadsheetApp.getUi();

    try {
        // 既存のトリガーを削除
        deleteTrigger();

        // 新しいトリガーを作成（毎日朝5時）
        ScriptApp.newTrigger('dailyBackupTrigger')
            .timeBased()
            .everyDays(1)
            .atHour(CONFIG.BACKUP_TIME_HOUR)
            .create();

        ui.alert('設定完了', `毎日${CONFIG.BACKUP_TIME_HOUR}時に自動バックアップが実行されます。`, ui.ButtonSet.OK);

    } catch (error) {
        console.error('トリガー設定エラー:', error);
        ui.alert('エラー', `トリガー設定に失敗しました：${error.message}`, ui.ButtonSet.OK);
    }
}

/**
 * 自動バックアップトリガーを削除
 */
function deleteTrigger() {
    const ui = SpreadsheetApp.getUi();

    try {
        const triggers = ScriptApp.getProjectTriggers();
        let deletedCount = 0;

        triggers.forEach(trigger => {
            if (trigger.getHandlerFunction() === 'dailyBackupTrigger') {
                ScriptApp.deleteTrigger(trigger);
                deletedCount++;
            }
        });

        if (deletedCount > 0) {
            ui.alert('完了', `${deletedCount}個の自動バックアップトリガーを削除しました。`, ui.ButtonSet.OK);
        } else {
            ui.alert('情報', '削除するトリガーが見つかりませんでした。', ui.ButtonSet.OK);
        }

    } catch (error) {
        console.error('トリガー削除エラー:', error);
        ui.alert('エラー', `トリガー削除に失敗しました：${error.message}`, ui.ButtonSet.OK);
    }
}

/**
 * トリガーから呼び出される日次バックアップ
 */
function dailyBackupTrigger() {
    try {
        const calendarIds = getStoredCalendarIds();
        if (calendarIds.length === 0) {
            console.log('設定されたカレンダーがありません。初期設定を実行してください。');
            return;
        }

        let totalEvents = 0;
        const errors = [];

        for (const calendarId of calendarIds) {
            try {
                const result = updateTodayAndFutureEvents(calendarId);
                if (result.success) {
                    totalEvents += result.eventCount;
                } else {
                    errors.push(`${calendarId}: ${result.error}`);
                }
            } catch (error) {
                errors.push(`${calendarId}: ${error.message}`);
            }
        }

        console.log(`自動バックアップ完了: ${totalEvents}件処理`);
        if (errors.length > 0) {
            console.error('バックアップエラー:', errors);
        }

    } catch (error) {
        console.error('自動バックアップエラー:', error);
    }
}

/**
 * カレンダー一覧を表示
 */
function showCalendarList() {
    const ui = SpreadsheetApp.getUi();

    try {
        const calendars = getCalendarList();
        const storedIds = getStoredCalendarIds();

        let message = '利用可能なカレンダー:\n\n';
        calendars.forEach((cal, index) => {
            const primaryMark = cal.isPrimary ? ' [プライマリ]' : '';
            const backupMark = storedIds.includes(cal.id) ? ' ✅' : '';
            message += `${index + 1}. ${cal.name}${primaryMark}${backupMark}\n`;
        });

        message += '\n✅ = バックアップ対象';

        ui.alert('カレンダー一覧', message, ui.ButtonSet.OK);

    } catch (error) {
        console.error('カレンダー一覧表示エラー:', error);
        ui.alert('エラー', `カレンダー一覧の取得に失敗しました：${error.message}`, ui.ButtonSet.OK);
    }
}

/**
 * 設定画面を表示
 */
function showSettings() {
    const ui = SpreadsheetApp.getUi();

    try {
        const storedIds = getStoredCalendarIds();
        const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyBackupTrigger');
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

        let message = '現在の設定:\n\n';
        message += `スプレッドシート: ${spreadsheet ? spreadsheet.getName() : '未設定'}\n`;
        message += `バックアップ対象カレンダー: ${storedIds.length}個\n`;
        message += `自動バックアップ: ${triggers.length > 0 ? '有効' : '無効'}\n`;
        message += `バックアップ時刻: 毎日${CONFIG.BACKUP_TIME_HOUR}時\n`;
        message += `バックアップ範囲: 今日から${CONFIG.BACKUP_MONTHS_FUTURE}ヶ月後まで\n`;
        message += `運用方法: 今日以降のデータ削除→再取得\n\n`;

        if (storedIds.length > 0) {
            message += 'バックアップ対象カレンダーID:\n';
            storedIds.forEach((id, index) => {
                message += `${index + 1}. ${id}\n`;
            });
        }

        ui.alert('設定情報', message, ui.ButtonSet.OK);

    } catch (error) {
        console.error('設定表示エラー:', error);
        ui.alert('エラー', `設定情報の取得に失敗しました：${error.message}`, ui.ButtonSet.OK);
    }
}

/**
 * 設定をリセット
 */
function resetSettings() {
    const ui = SpreadsheetApp.getUi();

    const response = ui.alert('確認', '全ての設定をリセットしますか？\n（バックアップデータは削除されません）', ui.ButtonSet.YES_NO);

    if (response === ui.Button.YES) {
        try {
            // トリガー削除
            deleteTrigger();

            // 保存されたカレンダーID削除
            PropertiesService.getScriptProperties().deleteProperty('CALENDAR_IDS');

            ui.alert('完了', '設定をリセットしました。', ui.ButtonSet.OK);

        } catch (error) {
            console.error('設定リセットエラー:', error);
            ui.alert('エラー', `設定リセットに失敗しました：${error.message}`, ui.ButtonSet.OK);
        }
    }
}