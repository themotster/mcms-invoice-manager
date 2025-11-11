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
  getClient: async (clientId) => await db.getClient(clientId),
  getClientByName: async (businessId, name) => await db.getClientByName(businessId, name),
  addClient: async (clientData) => await db.addClient(clientData),
  updateClient: async (clientId, clientData) => await db.updateClient(clientId, clientData),
  deleteClient: async (clientId) => await db.deleteClient(clientId),
  getClientDetails: async (clientId) => await db.getClientDetails(clientId),
  saveClientDetails: async (clientId, details) => await db.saveClientDetails(clientId, details || {}),
  getDocuments: async (options) => await db.getDocuments(options || {}),
  updateDocumentStatus: async (documentId, data) => await db.updateDocumentStatus(documentId, data || {}),
  getDocumentDefinitions: async (businessId, options) => await db.getDocumentDefinitions(businessId, options || {}),
  saveDocumentDefinition: async (businessId, definition) => await db.saveDocumentDefinition(businessId, definition),
  createNumberedDocument: async (options) => await documentService.createNumberedDocument(options || {}),
  createMCMSInvoice: async (options) => await documentService.createMCMSInvoice(options || {}),
  deleteDocument: async (documentId, options) => await documentService.deleteDocument(documentId, options || {}),
  getMergeFields: async () => await db.getMergeFields(),
  saveMergeField: async (field) => await db.saveMergeField(field || {}),
  getMaxInvoiceNumber: async (businessId) => await db.getMaxInvoiceNumber(businessId),
  setLastInvoiceNumber: async (businessId, val) => await db.setLastInvoiceNumber(businessId, val),
  documentNumberExists: async (businessId, docType, number) => await db.documentNumberExists(businessId, docType, number),
  filterDocumentsByExistingFiles: async (docs, options) => await documentService.filterDocumentsByExistingFiles(docs || [], options || {}),
  cleanOrphanDocuments: async (options) => await documentService.cleanOrphanDocuments(options || {}),
  indexInvoicesFromFilenames: async (options) => await documentService.indexInvoicesFromFilenames(options || {}),
  computeFinderInvoiceMax: async (options) => await documentService.computeFinderInvoiceMax(options || {}),
  getDocumentsByNumber: async (businessId, docType, number) => await db.getDocumentsByNumber(businessId, docType, number),
  scanTemplatePlaceholders: async (options) => await documentService.scanTemplatePlaceholders(options || {}),
  createNumberedWorkbookSimple: async (options) => await documentService.createNumberedWorkbookSimple(options || {}),
  chooseFile: async (options) => await ipcRenderer.invoke('choose-file', options || {}),
  chooseDirectory: async (options) => await ipcRenderer.invoke('choose-directory', options || {}),
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
  , businessSettings: async () => await db.businessSettings(),
  updateBusinessSettings: async (businessId, updates) => await db.updateBusinessSettings(businessId, updates || {})
  , deleteDocumentByPath: async (options) => await db.deleteDocumentByPath(options || {})
});
