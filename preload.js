const { contextBridge, ipcRenderer } = require('electron');
const path = require('path');
const db = require(path.join(__dirname, 'db.js'));
const documentService = require(path.join(__dirname, 'documentService.js'));
const ahmenCosting = require(path.join(__dirname, 'ahmenCosting.js'));

const jobsheetListeners = new Set();

ipcRenderer.on('jobsheet-change', (_event, payload) => {
  jobsheetListeners.forEach(callback => {
    try {
      callback(payload);
    } catch (err) {
      console.error('Jobsheet change listener error', err);
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
  getDocuments: async (options) => await db.getDocuments(options),
  addDocument: async (documentData) => await documentService.createDocument(documentData),
  createDocument: async (documentData) => await documentService.createDocument(documentData),
  updateDocumentStatus: async (documentId, data) => await db.updateDocumentStatus(documentId, data),
  getMusiciansForEvent: async (eventId) => await db.getMusiciansForEvent(eventId),
  addMusicianToEvent: async (eventId, musicianData) => await db.addMusicianToEvent(eventId, musicianData),
  updateMusicianPayment: async (musicianId, data) => await db.updateMusicianPayment(musicianId, data),
  deleteMusician: async (musicianId) => await db.deleteMusician(musicianId),
  getTimekeeperSessions: async (options) => await db.getTimekeeperSessions(options),
  importTimekeeperSession: async (sessionData) => await db.importTimekeeperSession(sessionData),
  markSessionExported: async (sessionId, exported) => await db.markSessionExported(sessionId, exported),
  deleteTimekeeperSession: async (sessionId, options) => await db.deleteTimekeeperSession(sessionId, options),
  deleteDocument: async (documentId, options) => await documentService.deleteDocument(documentId, options),
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
  getAhmenPricing: async () => await ahmenCosting.loadPricingConfig(),
  updateAhmenPricingService: async (serviceId, singers) => await ahmenCosting.savePricingServiceRoster(serviceId, singers),
  updateAhmenSingerPool: async (singers) => await ahmenCosting.saveSingerPool(singers),
  businessSettings: async () => await db.businessSettings(),
  openJobsheetWindow: async (options) => ipcRenderer.invoke('open-jobsheet-window', options || {}),
  notifyJobsheetChange: (payload) => ipcRenderer.send('jobsheet-change', payload || {}),
  onJobsheetChange: (callback) => {
    if (typeof callback !== 'function') return () => {};
    jobsheetListeners.add(callback);
    return () => jobsheetListeners.delete(callback);
  }
});
