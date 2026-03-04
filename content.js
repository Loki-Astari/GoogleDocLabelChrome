/**
 * Google Docs Labels - Chrome Extension (Manifest V3)
 * - Adds a Labels section to Google Docs left sidebar
 * - Adds a "Labels" item in Google Drive sidebar that opens an overlay
 *
 * Per-document labels are stored in the page's localStorage (docs.google.com origin).
 * Cross-domain master data + category config are stored in chrome.storage.local.
 */

(function() {
    'use strict';

    // -----------------------------
    // Extension storage (async)
    // -----------------------------

    const EXT_MASTER_KEY = 'gd-master-labels';       // { [labelName]: Array<{id,title,url}> }
    const EXT_CATEGORY_KEY = 'gd-label-categories';  // { assignments: { [labelName]: category }, categoryOrder: string[] }

    function storageGet(key, defaultValue) {
        return new Promise((resolve) => {
            try {
                chrome.storage.local.get({ [key]: defaultValue }, (result) => resolve(result[key]));
            } catch (e) {
                resolve(defaultValue);
            }
        });
    }

    function storageSet(key, value) {
        return new Promise((resolve) => {
            try {
                chrome.storage.local.set({ [key]: value }, () => resolve());
            } catch (e) {
                resolve();
            }
        });
    }

    async function loadCategoryConfig() {
        const config = await storageGet(EXT_CATEGORY_KEY, { assignments: {}, categoryOrder: [] });
        if (!config || typeof config !== 'object') return { assignments: {}, categoryOrder: [] };
        config.assignments = config.assignments || {};
        config.categoryOrder = config.categoryOrder || [];
        return config;
    }

    function saveCategoryConfig(config) {
        return storageSet(EXT_CATEGORY_KEY, config);
    }

    // -----------------------------
    // Shared helpers (Docs & Drive)
    // -----------------------------

    let labels = [];
    let labelsListContainer = null;
    let noLabelsMessage = null;
    let draggedIndex = null;
    let documentId = null;
    let expandedLabels = {};
    let lastKnownLabelsJson = '';

    function isGoogleDrive() {
        return window.location.hostname === 'drive.google.com';
    }

    function getDocumentId() {
        const match = window.location.pathname.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
        return match ? match[1] : null;
    }

    function getDocumentIdFromUrl(url) {
        const match = url.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
        return match ? match[1] : null;
    }

    function getDocumentTitle() {
        const titleElement = document.querySelector('.docs-title-input') ||
                            document.querySelector('[data-tooltip="Rename"]') ||
                            document.querySelector('.docs-title-widget');
        if (titleElement) {
            return titleElement.value || titleElement.textContent || 'Untitled';
        }
        const pageTitle = document.title.replace(' - Google Docs', '').trim();
        return pageTitle || 'Untitled';
    }

    function getStorageKey(docId) {
        return 'gd-labels-' + (docId || documentId);
    }

    // Incrementally update master label data for the current document only.
    async function updateMasterLabelList() {
        if (!documentId) return;
        try {
            let master = await storageGet(EXT_MASTER_KEY, {});
            // Handle legacy format (array)
            if (Array.isArray(master)) master = {};
            if (!master || typeof master !== 'object') master = {};

            const docTitle = getDocumentTitle();
            const docUrl = window.location.href;

            // Remove this doc from all labels
            Object.keys(master).forEach((label) => {
                master[label] = (master[label] || []).filter((doc) => doc && doc.id !== documentId);
                if (!master[label] || master[label].length === 0) delete master[label];
            });

            // Add current labels
            labels.forEach((label) => {
                if (!master[label]) master[label] = [];
                master[label].push({ id: documentId, title: docTitle, url: docUrl });
            });

            await storageSet(EXT_MASTER_KEY, master);
        } catch (e) {
            console.log('Google Docs Labels: Could not update master label list', e);
        }
    }

    // -----------------------------
    // Google Docs sidebar features
    // -----------------------------

    function saveLabels() {
        if (!documentId) return;
        try {
            const data = {
                labels: labels,
                title: getDocumentTitle(),
                url: window.location.href
            };
            localStorage.setItem(getStorageKey(), JSON.stringify(data));
            lastKnownLabelsJson = JSON.stringify(labels);
            void updateMasterLabelList();
        } catch (e) {
            console.log('Google Docs Labels: Could not save labels', e);
        }
    }

    function loadLabels() {
        if (!documentId) return;
        try {
            const saved = localStorage.getItem(getStorageKey());
            if (saved) {
                const data = JSON.parse(saved);
                if (Array.isArray(data)) {
                    labels = data;
                } else {
                    labels = data.labels || [];
                }
            } else {
                labels = [];
            }
            lastKnownLabelsJson = JSON.stringify(labels);
        } catch (e) {
            console.log('Google Docs Labels: Could not load labels', e);
            labels = [];
            lastKnownLabelsJson = '[]';
        }
    }

    function checkAndReloadLabels() {
        if (!documentId) return;
        try {
            const saved = localStorage.getItem(getStorageKey());
            let currentLabels = [];
            if (saved) {
                const data = JSON.parse(saved);
                currentLabels = Array.isArray(data) ? data : (data.labels || []);
            }
            const currentJson = JSON.stringify(currentLabels);
            if (currentJson !== lastKnownLabelsJson) {
                labels = currentLabels;
                lastKnownLabelsJson = currentJson;
                updateLabelsDisplay();
                void updateMasterLabelList();
            }
        } catch (e) {
            console.log('Google Docs Labels: Error checking for label changes', e);
        }
    }

    function findDocumentsWithLabel(labelName) {
        const documents = [];
        const currentDocId = documentId;

        for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (key && key.startsWith('gd-labels-')) {
                try {
                    const docId = key.replace('gd-labels-', '');
                    const saved = localStorage.getItem(key);
                    const data = JSON.parse(saved);

                    let docLabels = [];
                    let docTitle = 'Untitled';
                    let docUrl = 'https://docs.google.com/document/d/' + docId + '/edit';

                    if (Array.isArray(data)) {
                        docLabels = data;
                    } else {
                        docLabels = data.labels || [];
                        docTitle = data.title || 'Untitled';
                        docUrl = data.url || docUrl;
                    }

                    if (docLabels.includes(labelName)) {
                        documents.push({
                            id: docId,
                            title: docTitle,
                            url: docUrl,
                            isCurrent: docId === currentDocId
                        });
                    }
                } catch (e) {}
            }
        }

        return documents;
    }

    function exportLabel(labelName) {
        const documents = findDocumentsWithLabel(labelName);
        const exportData = {
            label: labelName,
            documents: documents.map(doc => ({ title: doc.title, url: doc.url }))
        };
        return JSON.stringify(exportData, null, 2);
    }

    function importLabel(jsonString) {
        try {
            const importData = JSON.parse(jsonString);
            if (!importData.label || !Array.isArray(importData.documents)) {
                return { success: false, message: 'Invalid JSON format. Expected { label, documents }' };
            }

            const labelName = importData.label;
            let importedCount = 0;

            importData.documents.forEach(doc => {
                if (!doc.url) return;
                const docId = getDocumentIdFromUrl(doc.url);
                if (!docId) return;

                const storageKey = getStorageKey(docId);
                let existingData = { labels: [], title: doc.title || 'Untitled', url: doc.url };

                try {
                    const saved = localStorage.getItem(storageKey);
                    if (saved) {
                        const data = JSON.parse(saved);
                        if (Array.isArray(data)) {
                            existingData.labels = data;
                        } else {
                            existingData = data;
                            existingData.labels = existingData.labels || [];
                        }
                    }
                } catch (e) {}

                if (!existingData.labels.includes(labelName)) {
                    existingData.labels.push(labelName);
                    localStorage.setItem(storageKey, JSON.stringify(existingData));
                    importedCount++;
                }
            });

            loadLabels();
            updateLabelsDisplay();
            void updateMasterLabelList();

            return {
                success: true,
                message: `Imported label "${labelName}" to ${importedCount} document(s).`
            };
        } catch (e) {
            return { success: false, message: 'Invalid JSON: ' + e.message };
        }
    }

    function showExportDialog(labelName) {
        const existingDialog = document.querySelector('#gd-label-dialog-overlay');
        if (existingDialog) existingDialog.remove();

        const jsonData = exportLabel(labelName);

        const overlay = document.createElement('div');
        overlay.id = 'gd-label-dialog-overlay';
        overlay.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.4); z-index: 10000; display: flex; align-items: center; justify-content: center;';

        const dialog = document.createElement('div');
        dialog.style.cssText = 'background: white; border-radius: 8px; padding: 24px; min-width: 400px; max-width: 520px; box-shadow: 0 4px 20px rgba(0,0,0,0.3);';

        const title = document.createElement('div');
        title.style.cssText = 'font-size: 16px; font-weight: 500; color: #202124; margin-bottom: 16px;';
        title.textContent = 'Export Label: ' + labelName;

        const textArea = document.createElement('textarea');
        textArea.value = jsonData;
        textArea.readOnly = true;
        textArea.style.cssText = 'width: 100%; height: 220px; padding: 10px 12px; border: 1px solid #dadce0; border-radius: 4px; font-size: 12px; font-family: monospace; box-sizing: border-box; resize: vertical;';

        const instructions = document.createElement('div');
        instructions.style.cssText = 'margin-top: 12px; padding: 12px; background: #f8f9fa; border-radius: 4px; font-size: 13px; color: #5f6368;';
        instructions.textContent = 'Copy the text above and send it to another user. They can import this label using the import button (↓) next to the Labels header.';

        const buttonContainer = document.createElement('div');
        buttonContainer.style.cssText = 'display: flex; justify-content: flex-end; gap: 12px; margin-top: 20px;';

        const copyBtn = document.createElement('button');
        copyBtn.textContent = 'Copy to Clipboard';
        copyBtn.style.cssText = 'padding: 8px 16px; border: none; background: transparent; color: #1a73e8; font-size: 14px; font-weight: 500; cursor: pointer; border-radius: 4px;';

        const okBtn = document.createElement('button');
        okBtn.textContent = 'OK';
        okBtn.style.cssText = 'padding: 8px 16px; border: none; background: #1a73e8; color: white; font-size: 14px; font-weight: 500; cursor: pointer; border-radius: 4px;';

        buttonContainer.appendChild(copyBtn);
        buttonContainer.appendChild(okBtn);
        dialog.appendChild(title);
        dialog.appendChild(textArea);
        dialog.appendChild(instructions);
        dialog.appendChild(buttonContainer);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        textArea.select();

        const closeDialog = () => overlay.remove();

        copyBtn.addEventListener('click', () => {
            textArea.select();
            navigator.clipboard.writeText(jsonData).then(() => {
                copyBtn.textContent = 'Copied!';
                setTimeout(() => { copyBtn.textContent = 'Copy to Clipboard'; }, 2000);
            }).catch(() => {
                try { document.execCommand('copy'); } catch (e) {}
                copyBtn.textContent = 'Copied!';
                setTimeout(() => { copyBtn.textContent = 'Copy to Clipboard'; }, 2000);
            });
        });

        okBtn.addEventListener('click', closeDialog);
        overlay.addEventListener('click', (e) => { if (e.target === overlay) closeDialog(); });
    }

    function showImportDialog() {
        const existingDialog = document.querySelector('#gd-label-dialog-overlay');
        if (existingDialog) existingDialog.remove();

        const overlay = document.createElement('div');
        overlay.id = 'gd-label-dialog-overlay';
        overlay.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.4); z-index: 10000; display: flex; align-items: center; justify-content: center;';

        const dialog = document.createElement('div');
        dialog.style.cssText = 'background: white; border-radius: 8px; padding: 24px; min-width: 400px; max-width: 520px; box-shadow: 0 4px 20px rgba(0,0,0,0.3);';

        const title = document.createElement('div');
        title.style.cssText = 'font-size: 16px; font-weight: 500; color: #202124; margin-bottom: 16px;';
        title.textContent = 'Import Label';

        const textArea = document.createElement('textarea');
        textArea.placeholder = 'Paste the exported JSON here...';
        textArea.style.cssText = 'width: 100%; height: 220px; padding: 10px 12px; border: 1px solid #dadce0; border-radius: 4px; font-size: 12px; font-family: monospace; box-sizing: border-box; resize: vertical;';

        const statusMsg = document.createElement('div');
        statusMsg.style.cssText = 'margin-top: 12px; padding: 12px; border-radius: 4px; font-size: 13px; display: none;';

        const instructions = document.createElement('div');
        instructions.style.cssText = 'margin-top: 12px; padding: 12px; background: #f8f9fa; border-radius: 4px; font-size: 13px; color: #5f6368;';
        instructions.textContent = 'Paste the JSON that was exported by another user. This will add the label to the documents in your localStorage.';

        const buttonContainer = document.createElement('div');
        buttonContainer.style.cssText = 'display: flex; justify-content: flex-end; gap: 12px; margin-top: 20px;';

        const cancelBtn = document.createElement('button');
        cancelBtn.textContent = 'Cancel';
        cancelBtn.style.cssText = 'padding: 8px 16px; border: none; background: transparent; color: #1a73e8; font-size: 14px; font-weight: 500; cursor: pointer; border-radius: 4px;';

        const importBtn = document.createElement('button');
        importBtn.textContent = 'Import';
        importBtn.style.cssText = 'padding: 8px 16px; border: none; background: #1a73e8; color: white; font-size: 14px; font-weight: 500; cursor: pointer; border-radius: 4px;';

        buttonContainer.appendChild(cancelBtn);
        buttonContainer.appendChild(importBtn);
        dialog.appendChild(title);
        dialog.appendChild(textArea);
        dialog.appendChild(statusMsg);
        dialog.appendChild(instructions);
        dialog.appendChild(buttonContainer);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        textArea.focus();

        const closeDialog = () => overlay.remove();
        cancelBtn.addEventListener('click', closeDialog);

        importBtn.addEventListener('click', () => {
            const result = importLabel(textArea.value);
            statusMsg.style.display = 'block';
            if (result.success) {
                statusMsg.style.background = '#e6f4ea';
                statusMsg.style.color = '#137333';
                statusMsg.textContent = result.message;
                setTimeout(closeDialog, 2000);
            } else {
                statusMsg.style.background = '#fce8e6';
                statusMsg.style.color = '#c5221f';
                statusMsg.textContent = result.message;
            }
        });

        overlay.addEventListener('click', (e) => { if (e.target === overlay) closeDialog(); });
    }

    function populateDocumentList(container, labelName) {
        while (container.firstChild) container.removeChild(container.firstChild);
        const documents = findDocumentsWithLabel(labelName);

        if (documents.length === 0) {
            const emptyMsg = document.createElement('div');
            emptyMsg.style.cssText = 'padding: 4px 8px; color: #5f6368; font-size: 12px; font-style: italic;';
            emptyMsg.textContent = 'No documents';
            container.appendChild(emptyMsg);
            return;
        }

        documents.forEach(doc => {
            const docItem = document.createElement('div');
            docItem.style.cssText = 'padding: 4px 8px; font-size: 12px; border-radius: 4px; transition: background-color 0.15s;';

            if (doc.isCurrent) {
                docItem.style.color = '#1a73e8';
                docItem.style.fontWeight = '500';
                docItem.textContent = doc.title + ' (current)';
            } else {
                const link = document.createElement('a');
                link.href = doc.url;
                link.textContent = doc.title;
                link.style.cssText = 'color: #202124; text-decoration: none;';
                link.addEventListener('mouseenter', () => { link.style.textDecoration = 'underline'; });
                link.addEventListener('mouseleave', () => { link.style.textDecoration = 'none'; });
                docItem.appendChild(link);
            }

            docItem.addEventListener('mouseenter', () => { docItem.style.backgroundColor = '#f1f3f4'; });
            docItem.addEventListener('mouseleave', () => { docItem.style.backgroundColor = 'transparent'; });
            container.appendChild(docItem);
        });
    }

    function updateLabelsDisplay() {
        if (!labelsListContainer || !noLabelsMessage) return;

        while (labelsListContainer.firstChild) labelsListContainer.removeChild(labelsListContainer.firstChild);

        if (labels.length === 0) {
            noLabelsMessage.style.display = 'block';
            return;
        }

        noLabelsMessage.style.display = 'none';

        labels.forEach((label, index) => {
            const labelContainer = document.createElement('div');
            labelContainer.className = 'gd-label-container';

            const labelItem = document.createElement('div');
            labelItem.style.cssText = 'padding: 4px 16px; color: #202124; font-size: 13px; cursor: grab; display: flex; align-items: center; justify-content: space-between; border-radius: 4px; transition: background-color 0.15s;';
            labelItem.draggable = true;
            labelItem.dataset.index = index;

            labelItem.addEventListener('dragstart', (e) => {
                draggedIndex = index;
                labelItem.style.opacity = '0.5';
                e.dataTransfer.effectAllowed = 'move';
            });

            labelItem.addEventListener('dragend', () => {
                labelItem.style.opacity = '1';
                draggedIndex = null;
                const items = labelsListContainer.querySelectorAll('[draggable="true"]');
                items.forEach(item => { item.style.borderTop = 'none'; item.style.borderBottom = 'none'; });
            });

            labelItem.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                const items = labelsListContainer.querySelectorAll('[draggable="true"]');
                items.forEach(item => { item.style.borderTop = 'none'; item.style.borderBottom = 'none'; });
                const targetIndex = parseInt(labelItem.dataset.index, 10);
                if (draggedIndex !== null && targetIndex !== draggedIndex) {
                    if (targetIndex < draggedIndex) labelItem.style.borderTop = '2px solid #1a73e8';
                    else labelItem.style.borderBottom = '2px solid #1a73e8';
                }
            });

            labelItem.addEventListener('drop', (e) => {
                e.preventDefault();
                const targetIndex = parseInt(labelItem.dataset.index, 10);
                if (draggedIndex !== null && targetIndex !== draggedIndex) {
                    const draggedLabel = labels[draggedIndex];
                    labels.splice(draggedIndex, 1);
                    labels.splice(targetIndex, 0, draggedLabel);
                    saveLabels();
                    updateLabelsDisplay();
                }
            });

            labelItem.addEventListener('mouseenter', () => { labelItem.style.backgroundColor = '#f1f3f4'; });
            labelItem.addEventListener('mouseleave', () => { labelItem.style.backgroundColor = 'transparent'; });

            const isExpanded = expandedLabels[label] || false;
            const expandBtn = document.createElement('span');
            expandBtn.style.cssText = 'color: #5f6368; cursor: pointer; margin-right: 8px; font-size: 10px; user-select: none; transition: transform 0.2s;';
            expandBtn.textContent = '▶';
            expandBtn.style.display = 'inline-block';
            expandBtn.style.transform = isExpanded ? 'rotate(90deg)' : 'rotate(0deg)';

            const dragHandle = document.createElement('span');
            dragHandle.style.cssText = 'color: #9aa0a6; margin-right: 8px; cursor: grab; font-size: 10px;';
            dragHandle.textContent = '⋮⋮';

            const labelText = document.createElement('span');
            labelText.style.cssText = 'flex: 1;';
            labelText.textContent = label;

            const exportBtn = document.createElement('span');
            exportBtn.style.cssText = 'color: #5f6368; cursor: pointer; padding: 2px 6px; font-size: 11px;';
            exportBtn.textContent = '↑';
            exportBtn.title = 'Export label';
            exportBtn.addEventListener('click', (e) => { e.stopPropagation(); showExportDialog(label); });

            const removeBtn = document.createElement('span');
            removeBtn.style.cssText = 'color: #5f6368; cursor: pointer; padding: 2px 6px; font-size: 11px;';
            removeBtn.textContent = '×';
            removeBtn.addEventListener('click', (e) => { e.stopPropagation(); labels.splice(index, 1); saveLabels(); updateLabelsDisplay(); });

            labelItem.appendChild(expandBtn);
            labelItem.appendChild(dragHandle);
            labelItem.appendChild(labelText);
            labelItem.appendChild(exportBtn);
            labelItem.appendChild(removeBtn);
            labelContainer.appendChild(labelItem);

            const docListContainer = document.createElement('div');
            docListContainer.style.cssText = 'padding-left: 32px; display: ' + (isExpanded ? 'block' : 'none') + ';';
            labelContainer.appendChild(docListContainer);

            if (isExpanded) populateDocumentList(docListContainer, label);

            expandBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                expandedLabels[label] = !expandedLabels[label];
                const nowExpanded = expandedLabels[label];
                expandBtn.style.transform = nowExpanded ? 'rotate(90deg)' : 'rotate(0deg)';
                docListContainer.style.display = nowExpanded ? 'block' : 'none';
                if (nowExpanded) populateDocumentList(docListContainer, label);
                else while (docListContainer.firstChild) docListContainer.removeChild(docListContainer.firstChild);
            });

            labelsListContainer.appendChild(labelContainer);
        });
    }

    function showAddLabelDialog() {
        const existingDialog = document.querySelector('#gd-label-dialog-overlay');
        if (existingDialog) existingDialog.remove();

        const overlay = document.createElement('div');
        overlay.id = 'gd-label-dialog-overlay';
        overlay.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.4); z-index: 10000; display: flex; align-items: center; justify-content: center;';

        const dialog = document.createElement('div');
        dialog.style.cssText = 'background: white; border-radius: 8px; padding: 24px; min-width: 300px; box-shadow: 0 4px 20px rgba(0,0,0,0.3);';

        const title = document.createElement('div');
        title.style.cssText = 'font-size: 16px; font-weight: 500; color: #202124; margin-bottom: 16px;';
        title.textContent = 'Add Label';

        const input = document.createElement('input');
        input.type = 'text';
        input.placeholder = 'Enter label name';
        input.style.cssText = 'width: 100%; padding: 10px 12px; border: 1px solid #dadce0; border-radius: 4px; font-size: 14px; box-sizing: border-box; outline: none;';

        const buttonContainer = document.createElement('div');
        buttonContainer.style.cssText = 'display: flex; justify-content: flex-end; gap: 12px; margin-top: 20px;';

        const cancelBtn = document.createElement('button');
        cancelBtn.textContent = 'Cancel';
        cancelBtn.style.cssText = 'padding: 8px 16px; border: none; background: transparent; color: #1a73e8; font-size: 14px; font-weight: 500; cursor: pointer; border-radius: 4px;';

        const addBtn = document.createElement('button');
        addBtn.textContent = 'Add';
        addBtn.style.cssText = 'padding: 8px 16px; border: none; background: #1a73e8; color: white; font-size: 14px; font-weight: 500; cursor: pointer; border-radius: 4px;';

        buttonContainer.appendChild(cancelBtn);
        buttonContainer.appendChild(addBtn);
        dialog.appendChild(title);
        dialog.appendChild(input);
        dialog.appendChild(buttonContainer);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        setTimeout(() => input.focus(), 100);
        const closeDialog = () => overlay.remove();
        cancelBtn.addEventListener('click', closeDialog);

        function doAdd() {
            if (input.value && input.value.trim()) {
                labels.push(input.value.trim());
                saveLabels();
                updateLabelsDisplay();
            }
            closeDialog();
        }

        addBtn.addEventListener('click', doAdd);
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') doAdd();
            else if (e.key === 'Escape') closeDialog();
        });
        overlay.addEventListener('click', (e) => { if (e.target === overlay) closeDialog(); });
    }

    function createLabelsSection(documentTabsSection) {
        const parentContainer = documentTabsSection.parentElement;
        if (!parentContainer) return;
        if (document.querySelector('#gd-labels-section')) return;

        const labelsSection = document.createElement('div');
        labelsSection.id = 'gd-labels-section';

        const computedStyle = window.getComputedStyle(documentTabsSection);
        labelsSection.style.cssText = `margin-bottom: ${computedStyle.marginBottom}; padding: ${computedStyle.padding}; background: #fff;`;

        const headerRow = document.createElement('div');
        headerRow.style.cssText = 'display: flex; align-items: center; justify-content: space-between; padding: 8px 16px;';

        const headerText = document.createElement('div');
        headerText.textContent = 'Labels';
        headerText.style.cssText = 'font-size: 11px; font-weight: 500; color: #5f6368; text-transform: uppercase; letter-spacing: 0.8px;';

        const buttonGroup = document.createElement('div');
        buttonGroup.style.cssText = 'display: flex; gap: 4px;';

        const importButton = document.createElement('button');
        importButton.textContent = '↓';
        importButton.style.cssText = 'border: none; background: transparent; color: #5f6368; font-size: 14px; cursor: pointer; padding: 0 4px; line-height: 1;';
        importButton.title = 'Import label';
        importButton.addEventListener('click', showImportDialog);

        const plusButton = document.createElement('button');
        plusButton.textContent = '+';
        plusButton.style.cssText = 'border: none; background: transparent; color: #5f6368; font-size: 18px; cursor: pointer; padding: 0 4px; line-height: 1;';
        plusButton.title = 'Add label';
        plusButton.addEventListener('click', showAddLabelDialog);

        buttonGroup.appendChild(importButton);
        buttonGroup.appendChild(plusButton);
        headerRow.appendChild(headerText);
        headerRow.appendChild(buttonGroup);
        labelsSection.appendChild(headerRow);

        noLabelsMessage = document.createElement('div');
        noLabelsMessage.textContent = 'No labels';
        noLabelsMessage.style.cssText = 'padding: 4px 16px 4px 32px; color: #5f6368; font-size: 12px; font-style: italic;';
        labelsSection.appendChild(noLabelsMessage);

        labelsListContainer = document.createElement('div');
        labelsListContainer.id = 'gd-labels-list';
        labelsListContainer.style.cssText = 'padding-left: 16px;';
        labelsSection.appendChild(labelsListContainer);

        parentContainer.insertBefore(labelsSection, documentTabsSection);

        loadLabels();
        void updateMasterLabelList();
        updateLabelsDisplay();

        document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') checkAndReloadLabels(); });
        window.addEventListener('focus', () => checkAndReloadLabels());
    }

    function initDocs() {
        documentId = getDocumentId();
        if (!documentId) return;

        const observer = new MutationObserver((mutations, obs) => {
            const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
            let node;
            while (node = walker.nextNode()) {
                if (node.textContent.trim() === 'Document tabs') {
                    let section = node.parentElement;
                    for (let i = 0; i < 10 && section; i++) {
                        if (section.querySelector && section.querySelector('[role="tree"], [role="button"]')) {
                            createLabelsSection(section);
                            obs.disconnect();
                            return;
                        }
                        section = section.parentElement;
                    }
                }
            }
        });

        observer.observe(document.body, { childList: true, subtree: true });
        setTimeout(() => observer.disconnect(), 30000);
    }

    // -----------------------------
    // Google Drive overlay + injection
    // -----------------------------

    async function showDriveLabelsOverlay() {
        const existing = document.querySelector('#gd-labels-overlay');
        if (existing) { existing.remove(); return; }

        let masterData = await storageGet(EXT_MASTER_KEY, {});
        if (Array.isArray(masterData)) {
            const converted = {};
            masterData.forEach((l) => { converted[l] = []; });
            masterData = converted;
        }
        const labelNames = Object.keys(masterData || {}).sort();
        const catConfig = await loadCategoryConfig();
        const expanded = {};

        // Cleanup stale assignments
        Object.keys(catConfig.assignments).forEach((l) => { if (!masterData[l]) delete catConfig.assignments[l]; });
        void saveCategoryConfig(catConfig);

        function getGrouped() {
            const groups = {};
            catConfig.categoryOrder.forEach((c) => { groups[c] = []; });
            if (!groups['Un-Categorized']) groups['Un-Categorized'] = [];
            labelNames.forEach((l) => {
                const cat = catConfig.assignments[l] || 'Un-Categorized';
                if (!groups[cat]) groups[cat] = [];
                groups[cat].push(l);
            });
            return groups;
        }

        function getOrder() {
            const order = catConfig.categoryOrder.filter((c) => c !== 'Un-Categorized');
            order.push('Un-Categorized');
            return order;
        }

        const overlay = document.createElement('div');
        overlay.id = 'gd-labels-overlay';
        overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:10000;display:flex;align-items:center;justify-content:center;';

        const panel = document.createElement('div');
        panel.style.cssText = 'background:#fff;border-radius:12px;min-width:480px;max-width:620px;max-height:85vh;display:flex;flex-direction:column;box-shadow:0 8px 32px rgba(0,0,0,0.3);';

        const header = document.createElement('div');
        header.style.cssText = 'display:flex;align-items:center;justify-content:space-between;padding:20px 24px 16px;border-bottom:1px solid #e8eaed;flex-shrink:0;';

        const title = document.createElement('h2');
        title.textContent = 'Labels';
        title.style.cssText = 'font-size:20px;font-weight:400;color:#202124;margin:0;font-family:"Google Sans",Roboto,sans-serif;';

        const headerBtns = document.createElement('div');
        headerBtns.style.cssText = 'display:flex;align-items:center;gap:8px;';

        const addCatBtn = document.createElement('button');
        addCatBtn.textContent = '+ Category';
        addCatBtn.style.cssText = 'border:1px solid #dadce0;background:#fff;color:#1a73e8;font-size:13px;font-weight:500;cursor:pointer;border-radius:4px;padding:6px 14px;';

        const closeBtn = document.createElement('button');
        closeBtn.textContent = '×';
        closeBtn.style.cssText = 'border:none;background:transparent;font-size:24px;cursor:pointer;color:#5f6368;padding:4px 8px;border-radius:50%;line-height:1;';

        headerBtns.appendChild(addCatBtn);
        headerBtns.appendChild(closeBtn);
        header.appendChild(title);
        header.appendChild(headerBtns);
        panel.appendChild(header);

        const bodyEl = document.createElement('div');
        bodyEl.style.cssText = 'overflow-y:auto;padding:8px 0;';
        panel.appendChild(bodyEl);
        overlay.appendChild(panel);

        // Custom mouse drag state
        let drag = { active: false, label: null, ghost: null, sourceEl: null };
        let catSections = []; // {el,name}

        function cleanupDrag() {
            if (drag.ghost && drag.ghost.parentNode) drag.ghost.parentNode.removeChild(drag.ghost);
            if (drag.sourceEl) drag.sourceEl.style.opacity = '1';
            catSections.forEach((c) => { c.el.style.boxShadow = 'none'; c.el.style.background = ''; });
            drag = { active: false, label: null, ghost: null, sourceEl: null };
        }

        function onMouseMove(e) {
            if (!drag.active) return;
            e.preventDefault();
            drag.ghost.style.left = (e.clientX + 12) + 'px';
            drag.ghost.style.top = (e.clientY - 14) + 'px';
            catSections.forEach((c) => {
                const r = c.el.getBoundingClientRect();
                const inside = e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom;
                c.el.style.boxShadow = inside ? '0 0 0 2px #1a73e8' : 'none';
                c.el.style.background = inside ? '#e8f0fe' : '';
            });
        }

        async function onMouseUp(e) {
            if (!drag.active) return;
            let targetCat = null;
            catSections.forEach((c) => {
                const r = c.el.getBoundingClientRect();
                if (e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom) targetCat = c.name;
            });

            if (targetCat !== null && drag.label) {
                if (targetCat === 'Un-Categorized') delete catConfig.assignments[drag.label];
                else catConfig.assignments[drag.label] = targetCat;
                await saveCategoryConfig(catConfig);
            }

            cleanupDrag();
            document.removeEventListener('mousemove', onMouseMove, true);
            document.removeEventListener('mouseup', onMouseUp, true);
            if (targetCat !== null) render();
        }

        function startDrag(e, labelName, itemEl) {
            e.preventDefault();
            drag.active = true;
            drag.label = labelName;
            drag.sourceEl = itemEl;
            itemEl.style.opacity = '0.4';

            const ghost = document.createElement('div');
            ghost.style.cssText = 'position:fixed;z-index:10002;padding:6px 14px;background:#fff;border:1px solid #dadce0;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,0.2);font-size:13px;color:#202124;pointer-events:none;white-space:nowrap;';
            ghost.textContent = labelName;
            ghost.style.left = (e.clientX + 12) + 'px';
            ghost.style.top = (e.clientY - 14) + 'px';
            document.body.appendChild(ghost);
            drag.ghost = ghost;

            document.addEventListener('mousemove', onMouseMove, true);
            document.addEventListener('mouseup', onMouseUp, true);
        }

        function render() {
            while (bodyEl.firstChild) bodyEl.removeChild(bodyEl.firstChild);
            catSections = [];

            if (labelNames.length === 0) {
                const empty = document.createElement('div');
                empty.textContent = 'No labels yet. Add labels to your Google Docs to see them here.';
                empty.style.cssText = 'color:#5f6368;font-style:italic;padding:24px;text-align:center;font-size:14px;';
                bodyEl.appendChild(empty);
            }

            const grouped = getGrouped();
            const order = getOrder();

            order.forEach((catName) => {
                const catLabels = grouped[catName] || [];
                const catSection = document.createElement('div');
                catSection.style.cssText = 'margin:4px 12px;border:1px solid #e8eaed;border-radius:8px;overflow:hidden;';
                catSections.push({ el: catSection, name: catName });

                const catHeader = document.createElement('div');
                catHeader.style.cssText = 'display:flex;align-items:center;padding:10px 14px;background:#f8f9fa;cursor:default;user-select:none;font-size:13px;font-weight:500;color:#5f6368;text-transform:uppercase;letter-spacing:0.5px;';

                const catNameEl = document.createElement('span');
                catNameEl.style.cssText = 'flex:1;';
                catNameEl.textContent = catName;

                const catCount = document.createElement('span');
                catCount.style.cssText = 'font-size:12px;font-weight:400;margin-right:4px;text-transform:none;letter-spacing:normal;';
                catCount.textContent = catLabels.length + (catLabels.length === 1 ? ' label' : ' labels');

                catHeader.appendChild(catNameEl);
                catHeader.appendChild(catCount);

                if (catName !== 'Un-Categorized') {
                    const delBtn = document.createElement('button');
                    delBtn.textContent = '×';
                    delBtn.title = 'Delete category';
                    delBtn.style.cssText = 'border:none;background:transparent;color:#5f6368;font-size:16px;cursor:pointer;padding:0 4px;border-radius:4px;line-height:1;margin-left:4px;';
                    delBtn.addEventListener('click', async (e) => {
                        e.stopPropagation();
                        catLabels.forEach((l) => { delete catConfig.assignments[l]; });
                        catConfig.categoryOrder = catConfig.categoryOrder.filter((c) => c !== catName);
                        await saveCategoryConfig(catConfig);
                        render();
                    });
                    catHeader.appendChild(delBtn);
                }

                catSection.appendChild(catHeader);

                const labelContainer = document.createElement('div');
                labelContainer.style.cssText = 'min-height:4px;';

                if (catLabels.length === 0) {
                    const emptyDrop = document.createElement('div');
                    emptyDrop.style.cssText = 'padding:10px 16px;color:#9aa0a6;font-size:13px;font-style:italic;text-align:center;';
                    emptyDrop.textContent = 'Drag labels here';
                    labelContainer.appendChild(emptyDrop);
                }

                catLabels.forEach((labelName) => {
                    const docs = masterData[labelName] || [];
                    const wrapper = document.createElement('div');

                    const item = document.createElement('div');
                    item.style.cssText = 'padding:8px 14px 8px 20px;font-size:14px;color:#202124;cursor:grab;display:flex;align-items:center;user-select:none;border-top:1px solid #f1f3f4;';

                    // Start drag on mousedown anywhere on item except nodrag elements
                    item.addEventListener('mousedown', (e) => {
                        const t = e.target;
                        if (t && t.dataset && t.dataset.nodrag) return;
                        e.stopPropagation();
                        startDrag(e, labelName, item);
                    });

                    const dragHandle = document.createElement('span');
                    dragHandle.textContent = '⋮⋮';
                    dragHandle.style.cssText = 'color:#bdc1c6;margin-right:10px;font-size:10px;cursor:grab;';

                    const expandIcon = document.createElement('span');
                    expandIcon.textContent = '▶';
                    expandIcon.dataset.nodrag = 'true';
                    expandIcon.style.cssText = 'font-size:10px;color:#5f6368;margin-right:8px;transition:transform 0.15s;display:inline-block;cursor:pointer;';
                    if (expanded[labelName]) expandIcon.style.transform = 'rotate(90deg)';

                    const labelText = document.createElement('span');
                    labelText.style.cssText = 'flex:1;';
                    labelText.textContent = labelName;

                    const docCount = document.createElement('span');
                    docCount.style.cssText = 'color:#5f6368;font-size:12px;margin-left:8px;';
                    docCount.textContent = docs.length + (docs.length === 1 ? ' doc' : ' docs');

                    item.appendChild(dragHandle);
                    item.appendChild(expandIcon);
                    item.appendChild(labelText);
                    item.appendChild(docCount);
                    wrapper.appendChild(item);

                    const docList = document.createElement('div');
                    docList.style.cssText = 'padding:2px 0 8px 54px;' + (expanded[labelName] ? 'display:block;' : 'display:none;');

                    if (docs.length === 0) {
                        const emptyMsg = document.createElement('div');
                        emptyMsg.textContent = 'No documents';
                        emptyMsg.style.cssText = 'color:#5f6368;font-size:13px;font-style:italic;padding:4px 0;';
                        docList.appendChild(emptyMsg);
                    } else {
                        docs.forEach((doc) => {
                            const link = document.createElement('a');
                            link.href = doc.url;
                            link.textContent = doc.title;
                            link.dataset.nodrag = 'true';
                            link.style.cssText = 'display:block;color:#1a73e8;text-decoration:none;padding:4px 0;font-size:13px;';
                            docList.appendChild(link);
                        });
                    }

                    wrapper.appendChild(docList);

                    expandIcon.addEventListener('click', (e) => {
                        e.stopPropagation();
                        expanded[labelName] = !expanded[labelName];
                        docList.style.display = expanded[labelName] ? 'block' : 'none';
                        expandIcon.style.transform = expanded[labelName] ? 'rotate(90deg)' : 'rotate(0deg)';
                    });

                    labelContainer.appendChild(wrapper);
                });

                catSection.appendChild(labelContainer);
                bodyEl.appendChild(catSection);
            });
        }

        function showNewCategoryDialog() {
            const dlgOverlay = document.createElement('div');
            dlgOverlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.4);z-index:10001;display:flex;align-items:center;justify-content:center;';

            const dlg = document.createElement('div');
            dlg.style.cssText = 'background:#fff;border-radius:8px;padding:24px;min-width:320px;box-shadow:0 4px 20px rgba(0,0,0,0.3);';

            const dlgTitle = document.createElement('h2');
            dlgTitle.textContent = 'New Category';
            dlgTitle.style.cssText = 'font-size:16px;font-weight:500;margin:0 0 16px;';

            const input = document.createElement('input');
            input.type = 'text';
            input.placeholder = 'Category name';
            input.style.cssText = 'width:100%;padding:10px 12px;border:1px solid #dadce0;border-radius:4px;font-size:14px;outline:none;box-sizing:border-box;';

            const dlgBtns = document.createElement('div');
            dlgBtns.style.cssText = 'display:flex;justify-content:flex-end;gap:12px;margin-top:20px;';

            const cancelBtn = document.createElement('button');
            cancelBtn.textContent = 'Cancel';
            cancelBtn.style.cssText = 'padding:8px 16px;border:none;background:transparent;color:#1a73e8;font-size:14px;font-weight:500;cursor:pointer;border-radius:4px;';

            const createBtn = document.createElement('button');
            createBtn.textContent = 'Create';
            createBtn.style.cssText = 'padding:8px 16px;border:none;background:#1a73e8;color:#fff;font-size:14px;font-weight:500;cursor:pointer;border-radius:4px;';

            dlgBtns.appendChild(cancelBtn);
            dlgBtns.appendChild(createBtn);
            dlg.appendChild(dlgTitle);
            dlg.appendChild(input);
            dlg.appendChild(dlgBtns);
            dlgOverlay.appendChild(dlg);
            document.body.appendChild(dlgOverlay);
            setTimeout(() => input.focus(), 50);

            const closeDlg = () => dlgOverlay.remove();
            cancelBtn.addEventListener('click', closeDlg);
            dlgOverlay.addEventListener('click', (e) => { if (e.target === dlgOverlay) closeDlg(); });

            async function doCreate() {
                const name = (input.value || '').trim();
                if (!name || name === 'Un-Categorized') return;
                if (catConfig.categoryOrder.includes(name)) return;
                catConfig.categoryOrder.push(name);
                await saveCategoryConfig(catConfig);
                closeDlg();
                render();
            }

            createBtn.addEventListener('click', doCreate);
            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') void doCreate();
                if (e.key === 'Escape') closeDlg();
            });
        }

        addCatBtn.addEventListener('click', showNewCategoryDialog);
        closeBtn.addEventListener('click', () => overlay.remove());
        overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
        document.addEventListener('keydown', function handler(e) {
            if (e.key === 'Escape') { overlay.remove(); document.removeEventListener('keydown', handler); }
        });

        render();
        document.body.appendChild(overlay);
    }

    function createDriveLabelItem(starredNavTreeHeader) {
        if (document.querySelector('#gd-drive-label-item')) return;

        const labelItem = starredNavTreeHeader.cloneNode(true);
        labelItem.id = 'gd-drive-label-item';

        const textNodes = [];
        const walker = document.createTreeWalker(labelItem, NodeFilter.SHOW_TEXT, null, false);
        let tNode;
        while (tNode = walker.nextNode()) {
            if (tNode.textContent.trim().length > 0) textNodes.push(tNode);
        }
        if (textNodes.length > 0) textNodes[0].textContent = 'Labels';

        const svgEl = labelItem.querySelector('svg');
        if (svgEl) {
            svgEl.innerHTML = '<path d="M17.63 5.84C17.27 5.33 16.67 5 16 5L5 5.01C3.9 5.01 3 5.9 3 7v10c0 1.1.9 1.99 2 1.99L16 19c.67 0 1.27-.33 1.63-.84L22 12l-4.37-6.16z" fill="currentColor"/>';
        }

        labelItem.querySelectorAll('[aria-selected="true"]').forEach(el => el.setAttribute('aria-selected', 'false'));

        labelItem.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            void showDriveLabelsOverlay();
        });

        starredNavTreeHeader.parentNode.insertBefore(labelItem, starredNavTreeHeader.nextSibling);
    }

    function findStarredNavTreeHeader() {
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
        let node;
        while (node = walker.nextNode()) {
            if (node.textContent.trim() === 'Starred') {
                let container = node.parentElement;
                while (container && container !== document.body) {
                    if (container.getAttribute('aria-labelledby') === 'navTreeHeader') return container;
                    container = container.parentElement;
                }
            }
        }
        return null;
    }

    function initDrive() {
        const observer = new MutationObserver(() => {
            if (!document.querySelector('#gd-drive-label-item')) {
                const starred = findStarredNavTreeHeader();
                if (starred) createDriveLabelItem(starred);
            }
        });
        observer.observe(document.body, { childList: true, subtree: true });
    }

    // -----------------------------
    // Entry point
    // -----------------------------

    if (isGoogleDrive()) {
        if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initDrive);
        else initDrive();
    } else {
        if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initDocs);
        else initDocs();
    }
})();

