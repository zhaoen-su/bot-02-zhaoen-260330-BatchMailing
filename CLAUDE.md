# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Google Apps Script mail merge tool (based on Martin Hawksey's Apache 2.0 licensed sample, modified by CAPSULE CXO TEAM). It reads recipient data from a Google Sheet, matches a Gmail draft by subject line, and sends personalized emails using `{{placeholder}}` template syntax. The UI is in Traditional Chinese (zh-TW).

## Architecture

The entire project is a single `code.js` file intended to be used inside Google Apps Script (not Node.js). Key functions:

- `onOpen()` — adds a "批次寄信" (Batch Mail) menu to the Google Sheets UI
- `sendEmails()` — main entry point: prompts for a draft subject line, fetches matching Gmail draft as template, iterates sheet rows, sends emails via `GmailApp.sendEmail`, records sent timestamps
- `getGmailTemplateFromDrafts_()` — finds a Gmail draft by subject, extracts body/attachments/inline images
- `fillInTemplateFromObject_()` — replaces `{{columnName}}` placeholders in the template with row data

Column constants `RECIPIENT_COL` ("收件人") and `EMAIL_SENT_COL` ("已寄出") must match the spreadsheet header names.

## Development Notes

- This runs in the Google Apps Script runtime (V8), not Node.js. There is no `package.json`, no build step, and no test framework.
- Global objects like `SpreadsheetApp`, `GmailApp`, `Browser`, and `MailApp` are provided by the Apps Script environment.
- To deploy: paste `code.js` into the Apps Script editor bound to the target Google Sheet, or use `clasp` to push.
- The `@OnlyCurrentDoc` annotation limits the script's OAuth scope to the bound spreadsheet.
