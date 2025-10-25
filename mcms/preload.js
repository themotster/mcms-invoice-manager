const { contextBridge, ipcRenderer } = require('electron');
const path = require('path');
const db = require(path.join(__dirname, '..', 'db.js'));
const documentService = require(path.join(__dirname, '..', 'documentService.js'));

contextBridge.exposeInMainWorld('api', {
  getClients: async () => await db.getClients(),
  getDocuments: async (options) => await db.getDocuments(options || {}),
  getDocumentDefinitions: async (businessId, options) => await db.getDocumentDefinitions(businessId, options || {}),
  saveDocumentDefinition: async (businessId, definition) => await db.saveDocumentDefinition(businessId, definition),
  createNumberedDocument: async (options) => await documentService.createNumberedDocument(options || {}),
  deleteDocument: async (documentId, options) => await documentService.deleteDocument(documentId, options || {}),
  chooseFile: async (options) => await ipcRenderer.invoke('choose-file', options || {}),
  openPath: async (targetPath) => await ipcRenderer.invoke('open-path', targetPath),
  showItemInFolder: async (targetPath) => await ipcRenderer.invoke('show-item-in-folder', targetPath)
});

