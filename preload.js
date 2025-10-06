const { contextBridge, ipcRenderer } = require('electron');
const path = require('path');
const db = require(path.join(__dirname, 'db.js'));
const documentService = require(path.join(__dirname, 'documentService.js'));
const ahmenCosting = require(path.join(__dirname, 'ahmenCosting.js'));

const jobsheetListeners = new Set();
const documentListeners = new Set();

ipcRenderer.on('jobsheet-change', (_event, payload) => {
  jobsheetListeners.forEach(callback => {
    try {
      callback(payload);
    } catch (err) {
      console.error('Jobsheet change listener error', err);
    }
  });
});

ipcRenderer.on('documents-change', (_event, payload) => {
  documentListeners.forEach(callback => {
    try {
      callback(payload);
    } catch (err) {
      console.error('Documents change listener error', err);
    }
  });
});

contextBridge.exposeInMainWorld('api', {
  getInvoices: async () => await db.getInvoices(),
  getStatus: async () => await db.getStatus(),
  getClients: async () => await db.getClients(),
  markPaid: async (num) => await db.markPaid(num),
  resetStatus: async (num) => await db.resetStatus(num),
  deleteInvoice: async (num) => await db.deleteInvoice(num),
  addInvoice: async (clientId, amount, dueDate) => {
    return await db.addInvoice(clientId, amount, dueDate);
  },
  addClient: async (clientData) => await db.addClient(clientData),
  updateClient: async (clientId, clientData) => await db.updateClient(clientId, clientData),
  getEvents: async (options) => await db.getEvents(options),
  addEvent: async (eventData) => await db.addEvent(eventData),
  updateEvent: async (eventId, eventData) => await db.updateEvent(eventId, eventData),
  getDocuments: async (options) => await db.getDocuments(options || {}),
  getDocumentDefinitions: async (businessId, options) => await db.getDocumentDefinitions(businessId, options || {}),
  getDocumentDefinition: async (businessId, identifier) => await db.getDocumentDefinition(businessId, identifier),
  saveDocumentDefinition: async (businessId, definition) => await db.saveDocumentDefinition(businessId, definition),
  deleteDocumentDefinition: async (businessId, identifier) => await db.deleteDocumentDefinition(businessId, identifier),
  reorderDocumentDefinitions: async (businessId, orderedKeys) => await db.reorderDocumentDefinitions(businessId, orderedKeys),
  addDocument: async (documentData) => await db.addDocument(documentData),
  createDocument: async (payload) => await documentService.createDocument(payload),
  exportWorkbookPdfs: async (options) => await documentService.exportWorkbookPdfs(options || {}),
  preflightPdfExport: async (options) => await documentService.preflightPdfExport(options || {}),
  indexInvoicesFromFilenames: async (options) => await documentService.indexInvoicesFromFilenames(options || {}),
  computeFinderInvoiceMax: async (options) => await documentService.computeFinderInvoiceMax(options || {}),
  syncJobsheetOutputs: async (options) => await documentService.syncJobsheetOutputs(options || {}),
  updateDocumentStatus: async (documentId, data) => await db.updateDocumentStatus(documentId, data),
  setDocumentNumber: async (documentId, newNumber) => await db.setDocumentNumber(documentId, newNumber),
  getMaxInvoiceNumber: async (businessId) => await db.getMaxInvoiceNumber(businessId),
  setLastInvoiceNumber: async (businessId, val) => await db.setLastInvoiceNumber(businessId, val),
  promotePdfToInvoice: async (documentId, options) => await db.promotePdfToInvoice(documentId, options || {}),
  setDocumentLock: async (documentId, locked) => await db.setDocumentLock(documentId, locked),
  getMusiciansForEvent: async (eventId) => await db.getMusiciansForEvent(eventId),
  addMusicianToEvent: async (eventId, musicianData) => await db.addMusicianToEvent(eventId, musicianData),
  updateMusicianPayment: async (musicianId, data) => await db.updateMusicianPayment(musicianId, data),
  deleteMusician: async (musicianId) => await db.deleteMusician(musicianId),
  getTimekeeperSessions: async (options) => await db.getTimekeeperSessions(options),
  importTimekeeperSession: async (sessionData) => await db.importTimekeeperSession(sessionData),
  markSessionExported: async (sessionId, exported) => await db.markSessionExported(sessionId, exported),
  deleteTimekeeperSession: async (sessionId, options) => await db.deleteTimekeeperSession(sessionId, options),
  deleteDocument: async (documentId, options) => await documentService.deleteDocument(documentId, options || {}),
  watchDocuments: async (options) => await ipcRenderer.invoke('watch-documents', options || {}),
  unwatchDocuments: async (options) => await ipcRenderer.invoke('unwatch-documents', options || {}),
  filterDocumentsByExistingFiles: async (documents, options) => await documentService.filterDocumentsByExistingFiles(documents || [], options || {}),
  listJobsheetDocuments: async (options) => await documentService.listJobsheetDocuments(options || {}),
  ensureJobsheetFolder: async (options) => await documentService.ensureJobsheetFolder(options || {}),
  getBookingPackPdfs: async (options) => await documentService.getBookingPackPdfs(options || {}),
  sendMailViaGraph: async (options) => await documentService.sendMailViaGraph(options || {}),
  listJobFolderFiles: async (options) => await documentService.listJobFolderFiles(options || {}),
  getMailTemplates: async (options) => await documentService.getMailTemplates(options || {}),
  saveMailTemplates: async (options) => await documentService.saveMailTemplates(options || {}),
  getDefaultMailTemplates: async (options) => await documentService.getDefaultMailTemplates(options || {}),
  getMailSignature: async (options) => await documentService.getMailSignature(options || {}),
  saveMailSignature: async (options) => await documentService.saveMailSignature(options || {}),
  extractJobsheetFromFolder: async (options) => await documentService.extractJobsheetDataFromFolder(options || {}),
  cleanOrphanDocuments: async (options) => await documentService.cleanOrphanDocuments(options || {}),
  deleteEvent: async (eventId) => await db.deleteEvent(eventId),
  deleteClient: async (clientId) => await db.deleteClient(clientId),
  getAhmenJobsheets: async (options) => await db.getAhmenJobsheets(options),
  getAhmenJobsheet: async (jobsheetId) => await db.getAhmenJobsheet(jobsheetId),
  addAhmenJobsheet: async (data) => await db.addAhmenJobsheet(data),
  updateAhmenJobsheet: async (jobsheetId, data) => await db.updateAhmenJobsheet(jobsheetId, data),
  updateAhmenJobsheetStatus: async (jobsheetId, status) => await db.updateAhmenJobsheetStatus(jobsheetId, status),
  deleteAhmenJobsheet: async (jobsheetId) => await db.deleteAhmenJobsheet(jobsheetId),
  getAhmenVenues: async (options) => await db.getAhmenVenues(options),
  saveAhmenVenue: async (data) => await db.saveAhmenVenue(data),
  deleteAhmenVenue: async (venueId) => await db.deleteAhmenVenue(venueId),
  getMergeFields: async () => await db.getMergeFields(),
  getMergeFieldValueSources: async (fieldKeys) => await db.getMergeFieldValueSources(fieldKeys),
  saveMergeField: async (field) => await db.saveMergeField(field),
  setMergeFieldValueSource: async (fieldKey, source) => await db.setMergeFieldValueSource(fieldKey, source),
  clearMergeFieldValueSource: async (fieldKey) => await db.clearMergeFieldValueSource(fieldKey),
  deleteMergeField: async (fieldKey) => await db.deleteMergeField(fieldKey),
  getAhmenPricing: async () => await ahmenCosting.loadPricingConfig(),
  logEmail: async (payload) => await db.logEmail(payload || {}),
  listEmailLog: async (filter) => await db.listEmailLog(filter || {}),
  deleteEmailLog: async (id) => await db.deleteEmailLog(id),
  updateAhmenPricingService: async (serviceId, singers) => await ahmenCosting.savePricingServiceRoster(serviceId, singers),
  updateAhmenSingerPool: async (singers) => await ahmenCosting.saveSingerPool(singers),
  businessSettings: async () => await db.businessSettings(),
  updateBusinessSettings: async (businessId, updates) => await db.updateBusinessSettings(businessId, updates),
  relocateBusinessDocuments: async (options) => await db.relocateBusinessDocuments(options || {}),
  listDocumentTree: async (options) => await db.listDocumentTree(options || {}),
  deleteDocumentFolder: async (options) => await db.deleteDocumentFolder(options || {}),
  deleteDocumentByPath: async (options) => await db.deleteDocumentByPath(options || {}),
  emptyDocumentsTrash: async (options) => await db.emptyDocumentsTrash(options || {}),
  chooseDirectory: async (options) => await ipcRenderer.invoke('choose-directory', options || {}),
  chooseFile: async (options) => await ipcRenderer.invoke('choose-file', options || {}),
  openPath: async (targetPath) => await ipcRenderer.invoke('open-path', targetPath),
  showItemInFolder: async (targetPath) => await ipcRenderer.invoke('show-item-in-folder', targetPath),
  openTemplate: async (options) => await ipcRenderer.invoke('open-template', options || {}),
  normalizeTemplate: async (options) => await ipcRenderer.invoke('normalize-template', options || {}),
  openJobsheetWindow: async (options) => ipcRenderer.invoke('open-jobsheet-window', options || {}),
  notifyJobsheetChange: (payload) => ipcRenderer.send('jobsheet-change', payload || {}),
  onJobsheetChange: (callback) => {
    if (typeof callback !== 'function') return () => {};
    jobsheetListeners.add(callback);
    return () => jobsheetListeners.delete(callback);
  },
  onDocumentsChange: (callback) => {
    if (typeof callback !== 'function') return () => {};
    documentListeners.add(callback);
    return () => documentListeners.delete(callback);
  }
});
