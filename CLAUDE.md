# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Chrome extension (Manifest V3) with two main parts:

### Part 1: Google Docs Labeling
When viewing a Google Doc, engineers can add labels to the document. Labels appear in the left sidebar (above "Document tabs") where users can:
- Add/remove labels on the current document
- Drag to reorder labels
- Expand a label to see all other documents sharing that label
- Import/export labels to share with other users

### Part 2: Google Drive Labels View
In Google Drive, a "Labels" item appears in the sidebar. Clicking it opens a dialog showing all labels organized by category, allowing users to:
- View all labels across all documents
- Organize labels into categories via drag-and-drop
- Create/delete categories

## Development

### Loading the Extension

1. Open `chrome://extensions/`
2. Enable Developer mode
3. Click "Load unpacked" and select this folder
4. Reload extension after code changes

No build step required - the extension loads directly from source files.

## Architecture

### Storage Model

Two storage layers:
- **localStorage** (docs.google.com origin): Per-document label data stored under `gd-labels-{docId}`
- **chrome.storage.local**: Cross-origin master index (`gd-master-labels`) and category config (`gd-label-categories`)

### content.js Structure

Single IIFE containing all functionality, split into sections:

1. **Extension storage helpers** (lines 16-50): Async wrappers for `chrome.storage.local`
2. **Shared helpers** (lines 52-122): Document ID extraction, master list sync
3. **Google Docs sidebar** (lines 124-714): Label management UI, drag-drop reordering, import/export dialogs
4. **Google Drive overlay** (lines 716-end): Category-based label organization, drag-drop between categories

### DOM Injection Strategy

- **Docs**: MutationObserver watches for "Document tabs" text node, then traverses up to find the section container and inserts Labels section above it
- **Drive**: Waits for sidebar DOM, finds "Starred" item by text content, injects "Labels" item after it

### Key Functions

- `updateMasterLabelList()`: Syncs current doc's labels to the master index
- `createLabelsSection()`: Builds and injects the Docs sidebar UI
- `showDriveLabelsOverlay()`: Creates the Drive overlay with category management
- `importLabel()`/`exportLabel()`: JSON-based label sharing between users
