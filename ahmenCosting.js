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

function slugifyName(name) {
  return (name || '')
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    || 'singer';
}

function ensureUniqueId(baseId, usedIds, index) {
  const sanitizedBase = baseId && baseId.trim() ? baseId.trim() : `singer-${index + 1}`;
  let candidate = sanitizedBase;
  let attempt = 1;
  while (usedIds.has(candidate)) {
    candidate = `${sanitizedBase}-${attempt}`;
    attempt += 1;
  }
  return candidate;
}

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

  const singerPool = buildSingerPool(mergedServiceTypes, overrides?.singerPool);

  cachedPricing = {
    ...base,
    serviceTypes: mergedServiceTypes,
    singerPool
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

function normalizePoolInput(singers, usedIds = new Set()) {
  if (!Array.isArray(singers)) throw new Error('Singer pool must be an array');
  return singers
    .map((singer, index) => {
      if (!singer) return null;
      const name = (singer.name || '').toString().trim();
      if (!name) return null;
      const rawId = singer.id != null ? String(singer.id) : '';
      let candidateId = rawId || slugifyName(name);
      let id = candidateId;
      if (!usedIds.has(candidateId)) {
        usedIds.add(candidateId);
      } else if (rawId && usedIds.has(rawId)) {
        id = rawId;
      } else {
        id = ensureUniqueId(candidateId, usedIds, index);
        usedIds.add(id);
      }

      const feeNumber = Number(singer.fee);
      const defaultCostNumber = singer.defaultCost != null ? Number(singer.defaultCost) : undefined;

      const serviceFees = {};
      if (singer.serviceFees && typeof singer.serviceFees === 'object') {
        Object.entries(singer.serviceFees).forEach(([serviceId, details]) => {
          if (!details) return;
          const feeValue = Number(details.fee);
          serviceFees[serviceId] = {
            fee: Number.isFinite(feeValue) && feeValue >= 0 ? feeValue : 0,
            defaultIncluded: Boolean(details.defaultIncluded)
          };
        });
      }

      return {
        id,
        name,
        fee: Number.isFinite(feeNumber) && feeNumber >= 0 ? feeNumber : 0,
        defaultIncluded: Boolean(singer.defaultIncluded),
        availability: singer.availability ? String(singer.availability) : '',
        comments: singer.comments ? String(singer.comments) : '',
        defaultCost: Number.isFinite(defaultCostNumber) ? defaultCostNumber : undefined,
        serviceFees
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

function buildSingerPool(serviceTypes, overrides) {
  const usedIds = new Set();
  const byName = new Map();

  serviceTypes.forEach(service => {
    (service.singers || []).forEach((singer) => {
      const name = (singer.name || '').toString().trim();
      if (!name) return;
      const key = name.toLowerCase();
      const serviceFeeNumber = Number(singer.fee);
      const normalizedServiceFee = Number.isFinite(serviceFeeNumber) && serviceFeeNumber >= 0 ? serviceFeeNumber : 0;
      const serviceDetails = {
        fee: normalizedServiceFee,
        defaultIncluded: Boolean(singer.defaultIncluded)
      };

      if (byName.has(key)) {
        const existing = byName.get(key);
        existing.serviceFees[service.id] = serviceDetails;
        existing.defaultIncluded = existing.defaultIncluded || Boolean(singer.defaultIncluded);
        if ((!Number.isFinite(existing.fee) || existing.fee === 0) && normalizedServiceFee > 0) {
          existing.fee = normalizedServiceFee;
        }
        if (!existing.comments && singer.comments) existing.comments = String(singer.comments);
        if (!existing.availability && singer.availability) existing.availability = String(singer.availability);
        if (existing.defaultCost === undefined && singer.defaultCost != null) {
          const costNumber = Number(singer.defaultCost);
          if (Number.isFinite(costNumber)) existing.defaultCost = costNumber;
        }
        return;
      }

      const baseId = singer.id != null ? String(singer.id) : slugifyName(name);
      const id = ensureUniqueId(baseId, usedIds, byName.size);
      usedIds.add(id);

      const comments = singer.comments ? String(singer.comments) : '';
      const availability = singer.availability ? String(singer.availability) : '';
      const costNumber = singer.defaultCost != null ? Number(singer.defaultCost) : undefined;

      byName.set(key, {
        id,
        name,
        fee: normalizedServiceFee,
        defaultIncluded: Boolean(singer.defaultIncluded),
        availability,
        comments,
        defaultCost: Number.isFinite(costNumber) ? costNumber : undefined,
        serviceFees: { [service.id]: serviceDetails }
      });
    });
  });

  const normalizedOverrides = Array.isArray(overrides) && overrides.length
    ? normalizePoolInput(overrides, usedIds)
    : [];

  normalizedOverrides.forEach(override => {
    const key = override.name.toLowerCase();
    const existing = byName.get(key);
    if (existing) {
      byName.set(key, {
        ...existing,
        ...override,
        serviceFees: {
          ...existing.serviceFees,
          ...override.serviceFees
        }
      });
    } else {
      byName.set(key, override);
    }
  });

  return Array.from(byName.values()).sort((a, b) => a.name.localeCompare(b.name));
}

async function saveSingerPool(singers) {
  const normalized = normalizePoolInput(singers);
  const overrides = readPricingOverrides();
  overrides.singerPool = normalized;
  writePricingOverrides(overrides);
  cachedPricing = null;
  return loadPricingConfig();
}

module.exports = {
  loadPricingConfig,
  savePricingServiceRoster,
  saveSingerPool
};
