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
  cachedPricing = parseCostingSheet(sheet);
  return cachedPricing;
}

module.exports = {
  loadPricingConfig
};
