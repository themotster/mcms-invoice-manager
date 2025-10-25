const { contextBridge, ipcRenderer } = require('electron');
const path = require('path');
const db = require(path.join(__dirname, '..', 'db.js'));
const documentService = require(path.join(__dirname, '..', 'documentService.js'));

const documentListeners = new Set();
ipcRenderer.on('documents-change', (_event, payload) => {
  documentListeners.forEach(cb => { try { cb(payload); } catch (_) {} });
});

contextBridge.exposeInMainWorld('api', {
  getClients: async () => await db.getClients(),
  getDocuments: async (options) => await db.getDocuments(options || {}),
  getDocumentDefinitions: async (businessId, options) => await db.getDocumentDefinitions(businessId, options || {}),
  saveDocumentDefinition: async (businessId, definition) => await db.saveDocumentDefinition(businessId, definition),
  createNumberedDocument: async (options) => await documentService.createNumberedDocument(options || {}),
  createMCMSInvoice: async (options) => await documentService.createMCMSInvoice(options || {}),
  deleteDocument: async (documentId, options) => await documentService.deleteDocument(documentId, options || {}),
  getMergeFields: async () => await db.getMergeFields(),
  saveMergeField: async (field) => await db.saveMergeField(field || {}),
  filterDocumentsByExistingFiles: async (docs, options) => await documentService.filterDocumentsByExistingFiles(docs || [], options || {}),
  cleanOrphanDocuments: async (options) => await documentService.cleanOrphanDocuments(options || {}),
  chooseFile: async (options) => await ipcRenderer.invoke('choose-file', options || {}),
  openPath: async (targetPath) => await ipcRenderer.invoke('open-path', targetPath),
  showItemInFolder: async (targetPath) => await ipcRenderer.invoke('show-item-in-folder', targetPath),
  copyTextToClipboard: async (text) => await ipcRenderer.invoke('copy-text-to-clipboard', text || ''),
  watchDocuments: async (options) => await ipcRenderer.invoke('watch-documents', options || {}),
  unwatchDocuments: async (options) => await ipcRenderer.invoke('unwatch-documents', options || {}),
  onDocumentsChange: (callback) => {
    if (typeof callback !== 'function') return () => {};
    documentListeners.add(callback);
    return () => documentListeners.delete(callback);
  }
});
