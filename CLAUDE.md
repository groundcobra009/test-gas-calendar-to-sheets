# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script (GAS) project that backs up Google Calendar events to Google Sheets. The system provides automated daily backups with safety measures to prevent accidental operations on primary calendars.

## Architecture

### Core Components

1. **Calendar Backup System** (Code.js)
   - `performInitialBackup()`: Full backup (1 year past to 1 year future)
   - `performDailyBackup()`: Incremental backup (today onwards)
   - `backupCalendarBatch()`: Batch processing for large calendars
   - Safety requirement: Calendar ID must be explicitly specified

2. **Scheduling System**
   - Automatic daily triggers at 5 AM JST
   - `setupDailyTrigger()`: Creates time-based triggers
   - `removeDailyTrigger()`: Removes existing triggers

3. **User Interface**
   - Custom menu in Google Sheets
   - Confirmation dialogs for all operations
   - Progress tracking for long-running operations

### Key Configuration Constants

```javascript
const BACKUP_MONTHS_FUTURE = 12;        // Future months to backup
const BACKUP_TIME_HOUR = 5;             // Daily backup hour (5 AM)
const SPREADSHEET_NAME = 'カレンダーバックアップ';
const MAX_EVENTS_PER_BATCH = 100;       // Events per batch
const EXECUTION_TIME_LIMIT = 5 * 60 * 1000; // 5 minutes
```

## Development Commands

### Deployment
```bash
# Deploy via Google Apps Script IDE (no CLI deployment)
# Open script.google.com and use the Deploy button
```

### Testing
```javascript
// Test functions directly in GAS editor
// Use Run button in script.google.com
// Check execution logs via View > Logs
```

## Important Implementation Details

1. **Calendar ID Safety**: The system requires explicit calendar ID specification and shows warnings before operations
2. **Batch Processing**: Handles large calendars by processing in batches to avoid execution time limits
3. **State Management**: Uses Script Properties to track batch processing state
4. **Event Deduplication**: Uses event IDs to prevent duplicate entries
5. **Time Zone**: All operations use Asia/Tokyo timezone

## OAuth Scopes Required

The project requires these permissions (defined in appsscript.json):
- https://www.googleapis.com/auth/calendar.readonly
- https://www.googleapis.com/auth/spreadsheets
- https://www.googleapis.com/auth/drive
- https://www.googleapis.com/auth/script.scriptapp
- https://www.googleapis.com/auth/script.external_request