# Google Docs Labels - Chrome Extension

Adds a Labels section to the Google Docs left sidebar for organizing and categorizing documents.

## Features

- Add labels to any Google Doc
- Drag and drop to reorder labels
- Expand labels to see all documents with that label
- Export/import labels to share with other users
- Labels persist in localStorage per document
- Auto-reload when switching tabs/windows

## Installation

### Developer Mode (Unpacked Extension)

1. Open Chrome and navigate to `chrome://extensions/`
2. Enable **Developer mode** (toggle in the top right)
3. Click **Load unpacked**
4. Select the `chrome-extension` folder
5. The extension is now installed and active

### Usage

1. Open any Google Doc
2. Look for the **Labels** section in the left sidebar (above "Document tabs")
3. Click **+** to add a new label
4. Click **↓** to import a label from another user
5. Click **↑** on a label to export it
6. Click **▶** to expand a label and see all documents with that label
7. Drag labels to reorder them
8. Click **×** to remove a label

## Files

- `manifest.json` - Chrome extension manifest (Manifest V3)
- `content.js` - Content script injected into Google Docs pages

