const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const COSTING_CONFIG = {
  worksheet: 'Costing',
  templatePath: path.resolve(__dirname, 'AhMen Client Data and Docs Template.xlsx'),
  serviceTypes: [
    {
      id: 'service',
      label: 'Service',
      columns: { name: 'B', availability: 'C', fee: 'D', cost: 'E', comments: 'F' },
      totalCell: null
    },
    {
      id: 'dinner',
      label: 'Dinner',
      columns: { name: 'G', availability: 'H', fee: 'I', cost: 'J', comments: 'K' },
      totalCell: null
    },
    {
      id: 'service_plus_dinner',
      label: 'Service + Dinner',
      columns: { name: 'M', availability: 'N', fee: 'O', cost: 'P', comments: null },
      totalCell: null
    }
  ],
  dataRowStart: 6,
  dataRowEnd: 20
};

const PRICING_OVERRIDE_PATH = path.resolve(__dirname, 'ahmenPricingOverrides.json');

function readPricingOverrides() {
  try {
    const raw = fs.readFileSync(PRICING_OVERRIDE_PATH, 'utf-8');
    const data = JSON.parse(raw);
    if (data && typeof data === 'object') return data;
    return {};
  } catch (err) {
    if (err.code === 'ENOENT') return {};
    console.warn('Failed to read pricing overrides', err);
    return {};
  }
}

function writePricingOverrides(overrides) {
  try {
    fs.writeFileSync(PRICING_OVERRIDE_PATH, JSON.stringify(overrides, null, 2), 'utf-8');
  } catch (err) {
    console.error('Unable to persist pricing overrides', err);
    throw err;
  }
}

let cachedPricing = null;

function decodeValue(cell) {
  if (!cell) return '';
  let { value } = cell;
  if (value && typeof value === 'object') {
    if (value.result !== undefined) return value.result;
    if (value.text) return value.text;
    if (value.richText) return value.richText.map(rt => rt.text).join('');
    if (value.formula) return value.result;
    return '';
  }
  return value;
}

function parseCostingSheet(sheet) {
  const serviceTypes = COSTING_CONFIG.serviceTypes.map(typeConfig => {
    const singers = [];
    for (let row = COSTING_CONFIG.dataRowStart; row <= COSTING_CONFIG.dataRowEnd; row++) {
      const nameCell = sheet.getCell(`${typeConfig.columns.name}${row}`);
      const name = (decodeValue(nameCell) || '').toString().trim();
      if (!name) continue;
      if (name.toLowerCase() === 'quote') continue;

      const availability = (decodeValue(sheet.getCell(`${typeConfig.columns.availability}${row}`)) || '').toString().trim();
      const fee = Number(decodeValue(sheet.getCell(`${typeConfig.columns.fee}${row}`))) || 0;
      const cost = Number(decodeValue(sheet.getCell(`${typeConfig.columns.cost}${row}`))) || 0;
      const comments = typeConfig.columns.comments
        ? (decodeValue(sheet.getCell(`${typeConfig.columns.comments}${row}`)) || '').toString().trim()
        : '';

      singers.push({
        id: `${typeConfig.id}__${name.replace(/\s+/g, '_').toLowerCase()}`,
        name,
        fee,
        defaultIncluded: availability.toLowerCase() === 'yes',
        availability: availability || '',
        comments,
        defaultCost: cost
      });
    }

    return {
      id: typeConfig.id,
      label: typeConfig.label,
      totalSuggested: 0,
      singers
    };
  });

  return {
    serviceTypes,
    updatedAt: new Date().toISOString()
  };
}

async function loadPricingConfig() {
  if (cachedPricing) return cachedPricing;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(COSTING_CONFIG.templatePath);
  const sheet = workbook.getWorksheet(COSTING_CONFIG.worksheet);
  if (!sheet) throw new Error('Costing sheet not found in AhMen template');
  const base = parseCostingSheet(sheet);
  const overrides = readPricingOverrides();

  const mergedServiceTypes = base.serviceTypes.map(service => {
    const overrideList = overrides?.[service.id];
    if (Array.isArray(overrideList) && overrideList.length) {
      const normalized = overrideList.map((singer, index) => ({
        id: String(singer.id ?? `${service.id}-override-${index}`),
        name: singer.name || '',
        fee: singer.fee != null ? Number(singer.fee) : 0,
        defaultIncluded: Boolean(singer.defaultIncluded),
        availability: singer.availability || '',
        comments: singer.comments || '',
        defaultCost: singer.defaultCost != null ? Number(singer.defaultCost) : undefined
      }));
      return {
        ...service,
        singers: normalized
      };
    }
    return service;
  });

  cachedPricing = {
    ...base,
    serviceTypes: mergedServiceTypes
  };
  return cachedPricing;
}

function normalizeRosterInput(singers) {
  if (!Array.isArray(singers)) throw new Error('Singer list must be an array');
  return singers
    .map((singer, index) => {
      if (!singer) return null;
      const id = singer.id != null ? String(singer.id) : `custom-${index}`;
      const name = (singer.name || '').toString().trim();
      if (!name) return null;
      const feeNumber = Number(singer.fee);
      const fee = Number.isFinite(feeNumber) ? feeNumber : 0;
      return {
        id,
        name,
        fee,
        defaultIncluded: Boolean(singer.defaultIncluded),
        availability: singer.availability ? String(singer.availability) : '',
        comments: singer.comments ? String(singer.comments) : '',
        defaultCost: singer.defaultCost != null ? Number(singer.defaultCost) : undefined
      };
    })
    .filter(Boolean);
}

async function savePricingServiceRoster(serviceId, singers) {
  if (!serviceId) throw new Error('Service id is required');
  const normalized = normalizeRosterInput(singers);
  const overrides = readPricingOverrides();
  overrides[serviceId] = normalized;
  writePricingOverrides(overrides);
  cachedPricing = null;
  return loadPricingConfig();
}

module.exports = {
  loadPricingConfig,
  savePricingServiceRoster
};
