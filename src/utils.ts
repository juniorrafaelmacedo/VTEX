/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import * as XLSX from 'xlsx';
import { SpreadsheetMapping, VtexShippingRow, ConversionStats, WeightRange } from './types';

// Standard requested Intelipost weight list in KG
export const DEFAULT_WEIGHTS_KG = [
  0.25, 0.5, 0.75, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15, 20, 25, 30, 50, 100, 200
];

/**
 * Normalizes and formats CEP to the VTEX format: "00000-000"
 */
export function formatCep(val: any): string {
  if (val === null || val === undefined) return '';
  const digits = String(val).replace(/\D/g, '').padStart(8, '0');
  if (digits.length === 8) {
    return `${digits.slice(0, 5)}-${digits.slice(5)}`;
  }
  return digits;
}

/**
 * Auto-detects the header row, sheets, and default column configurations of an Intelipost spreadsheet.
 */
export function detectSpreadsheetDetails(
  sheetNames: string[],
  workbook: XLSX.WorkBook
): { suggestedSheet: string; suggestedMapping: SpreadsheetMapping; availableSheets: string[] } {
  // 1. Select sheet: Prefer sheets with '2.5' or just the first sheet.
  let suggestedSheet = sheetNames[0];
  const preferredSheet = sheetNames.find(name => name.includes('2.5'));
  if (preferredSheet) {
    suggestedSheet = preferredSheet;
  }

  const sheet = workbook.Sheets[suggestedSheet];
  const rows = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1 });

  // 2. Discover header row: scan first 15 rows for the best match.
  let bestRowIndex = 3; // Fallback to 3 (which is index 3 or row 4)
  let bestScore = -1;

  for (let rIdx = 0; rIdx < Math.min(rows.length, 15); rIdx++) {
    const row = rows[rIdx];
    if (!Array.isArray(row)) continue;

    let score = 0;
    row.forEach((cell) => {
      const valStr = String(cell || '').toLowerCase().trim();
      if (valStr.includes('cep') || valStr.includes('inicial') || valStr.includes('zip')) score += 5;
      if (valStr.includes('final') || valStr.includes('fim')) score += 5;
      if (valStr.includes('prazo') || valStr.includes('entrega') || valStr.includes('dias')) score += 4;
      if (valStr.includes('excedente') || valStr.includes('extra') || valStr.includes('adicional')) score += 3;
      if (valStr.includes('gris')) score += 3;
      if (valStr.includes('seguro') || valStr.includes('ad valorem') || valStr.includes('advalorem')) score += 3;
      
      // Check if item looks like numeric weight
      const parsed = parseFloat(valStr);
      if (!isNaN(parsed) && parsed > 0 && parsed <= 500) {
        score += 2;
      }
    });

    if (score > bestScore) {
      bestScore = score;
      bestRowIndex = rIdx;
    }
  }

  const headerList = rows[bestRowIndex] || [];

  // 3. Match specific columns
  let zipCodeStartCol = 8; // standard index fallback
  let zipCodeEndCol = 9;   // standard index fallback
  let prazoCol = 10;       // standard index fallback
  let excedenteCol = 30;   // standard index fallback
  let grisCol: number | null = null;
  let seguroCol: number | null = null;

  headerList.forEach((cell, index) => {
    const valStr = String(cell || '').toLowerCase().trim();
    
    // CEP Inicial: can have keywords like 'cep inicial', 'cep de início', 'zipcode start'
    if (valStr.includes('cep') && (valStr.includes('inicial') || valStr.includes('inicio') || valStr.includes('início') || valStr.includes('start'))) {
      zipCodeStartCol = index;
    }
    // CEP Final: can have keywords like 'cep final', 'cep de fim', 'zipcode end'
    else if (valStr.includes('cep') && (valStr.includes('final') || valStr.includes('fim') || valStr.includes('end'))) {
      zipCodeEndCol = index;
    }
    // Prazo
    else if (valStr === 'prazo' || valStr.includes('prazo de entrega') || valStr.includes('tempo de entrega') || valStr.includes('timecost')) {
      prazoCol = index;
    }
    // Excedente
    else if (valStr.includes('excedente') || valStr.includes('extra') || valStr.includes('adicional') || valStr.includes('adicional por kg')) {
      excedenteCol = index;
    }
    // GRIS
    else if (valStr.includes('gris')) {
      grisCol = index;
    }
    // Seguro / Ad Valorem
    else if (valStr.includes('seguro') || valStr.includes('advalorem') || valStr.includes('ad valorem') || valStr.includes('ad-valorem')) {
      seguroCol = index;
    }
  });

  // Try to find the weights within the header
  const weightCols: number[] = [];
  DEFAULT_WEIGHTS_KG.forEach((wKg) => {
    // Exact or close match in numeric headers
    for (let cIdx = 0; cIdx < headerList.length; cIdx++) {
      const cell = headerList[cIdx];
      const parsed = parseFloat(String(cell || '').replace(',', '.'));
      if (!isNaN(parsed) && Math.abs(parsed - wKg) < 0.001) {
        weightCols.push(cIdx);
        break;
      }
    }
  });

  // If we couldn't match weight columns automatically, scan any columns with weight headers
  if (weightCols.length === 0) {
    headerList.forEach((cell, index) => {
      const parsed = parseFloat(String(cell || '').replace(',', '.'));
      if (!isNaN(parsed) && parsed > 0 && parsed <= 500) {
        weightCols.push(index);
      }
    });
    // Sort columns by numerical value in headers to be logical
    weightCols.sort((a, b) => {
      const valA = parseFloat(String(headerList[a]).replace(',', '.'));
      const valB = parseFloat(String(headerList[b]).replace(',', '.'));
      return valA - valB;
    });
  }

  return {
    suggestedSheet,
    suggestedMapping: {
      sheetName: suggestedSheet,
      headerRowIndex: bestRowIndex,
      zipCodeStartCol,
      zipCodeEndCol,
      prazoCol,
      excedenteCol,
      grisCol,
      seguroCol,
      weightCols,
    },
    availableSheets: sheetNames,
  };
}

/**
 * Main conversion routine
 */
export function processIntelipostToVtex(
  workbook: XLSX.WorkBook,
  mapping: SpreadsheetMapping,
  weightsConfig: number[]
): { vtexRows: VtexShippingRow[]; stats: ConversionStats } {
  const sheet = workbook.Sheets[mapping.sheetName];
  if (!sheet) {
    throw new Error(`Planilha "${mapping.sheetName}" não foi encontrada no arquivo.`);
  }

  const rows = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1 });
  const headerList = rows[mapping.headerRowIndex] || [];

  // Map user weights or auto-detected weights to column indexes
  const weightRanges: WeightRange[] = [];
  let lastWeightEndGrams = 0;

  // Let's match requested weights with column index
  weightsConfig.forEach((wKg) => {
    // Find matching column index
    let matchedColIdx = -1;
    for (let cIdx = 0; cIdx < headerList.length; cIdx++) {
      const cell = headerList[cIdx];
      const parsed = parseFloat(String(cell || '').replace(',', '.'));
      if (!isNaN(parsed) && Math.abs(parsed - wKg) < 0.001) {
        matchedColIdx = cIdx;
        break;
      }
    }

    if (matchedColIdx !== -1) {
      const currentEndGrams = Math.round(wKg * 1000);
      const startGrams = lastWeightEndGrams === 0 ? 0 : lastWeightEndGrams + 1;
      
      weightRanges.push({
        colIdx: matchedColIdx,
        label: `${wKg} KG`,
        weightKg: wKg,
        start: startGrams,
        end: currentEndGrams,
      });

      lastWeightEndGrams = currentEndGrams;
    }
  });

  // If no columns perfectly match the selected weights, fallback to weightCols already detected
  if (weightRanges.length === 0 && mapping.weightCols.length > 0) {
    mapping.weightCols.forEach((colIdx) => {
      const val = parseFloat(String(headerList[colIdx] || '').replace(',', '.'));
      if (!isNaN(val)) {
        const currentEndGrams = Math.round(val * 1000);
        const startGrams = lastWeightEndGrams === 0 ? 0 : lastWeightEndGrams + 1;
        
        weightRanges.push({
          colIdx,
          label: `${val} KG`,
          weightKg: val,
          start: startGrams,
          end: currentEndGrams,
        });
        
        lastWeightEndGrams = currentEndGrams;
      }
    });
  }

  const vtexRows: VtexShippingRow[] = [];
  const startRow = mapping.headerRowIndex + 1;
  const uniqueZips = new Set<string>();
  let totalCostSum = 0;
  let maxWeightKg = 0;
  let minTime = Infinity;
  let maxTime = -Infinity;

  let outRowCounter = 1;

  for (let rIdx = startRow; rIdx < rows.length; rIdx++) {
    const row = rows[rIdx];
    if (!row || row.length === 0) continue;

    const zipStartRaw = row[mapping.zipCodeStartCol];
    const zipEndRaw = row[mapping.zipCodeEndCol];

    // Skip empty lines or rows where the zip codes are invalid / empty
    if (zipStartRaw === null || zipStartRaw === undefined || String(zipStartRaw).trim() === '') continue;
    if (zipEndRaw === null || zipEndRaw === undefined || String(zipEndRaw).trim() === '') continue;

    const zipStart = formatCep(zipStartRaw);
    const zipEnd = formatCep(zipEndRaw);
    
    // Add to unique zip track
    uniqueZips.add(zipStart);
    uniqueZips.add(zipEnd);

    const rawPrazo = parseInt(String(row[mapping.prazoCol] || '0').replace(/\D/g, ''), 10);
    const prazo = isNaN(rawPrazo) ? 1 : rawPrazo;

    if (prazo < minTime) minTime = prazo;
    if (prazo > maxTime) maxTime = prazo;

    const rawExcedente = parseFloat(String(row[mapping.excedenteCol] || '0').replace(',', '.'));
    const valorExcedente = isNaN(rawExcedente) ? 0 : rawExcedente;

    // Price percent (Gris + Seguro)
    let grisValue = 0;
    if (mapping.grisCol !== null && mapping.grisCol !== undefined) {
      const parsedGris = parseFloat(String(row[mapping.grisCol] || '0').replace(',', '.'));
      grisValue = isNaN(parsedGris) ? 0 : parsedGris;
    }

    let seguroValue = 0;
    if (mapping.seguroCol !== null && mapping.seguroCol !== undefined) {
      const parsedSeguro = parseFloat(String(row[mapping.seguroCol] || '0').replace(',', '.'));
      seguroValue = isNaN(parsedSeguro) ? 0 : parsedSeguro;
    }

    const pricePercent = (grisValue + seguroValue);

    // Iterate through weights and append VTEX row
    weightRanges.forEach((r) => {
      const rawCost = row[r.colIdx];
      if (rawCost === null || rawCost === undefined || String(rawCost).trim() === '') return;

      const cost = parseFloat(String(rawCost).replace(',', '.'));
      if (isNaN(cost)) return;

      totalCostSum += cost;
      if (r.weightKg > maxWeightKg) {
        maxWeightKg = r.weightKg;
      }

      vtexRows.push({
        index: outRowCounter++,
        zipCodeStart: zipStart,
        zipCodeEnd: zipEnd,
        polygonName: '',
        weightStart: r.start,
        weightEnd: r.end,
        absoluteMoneyCost: cost,
        pricePercent: pricePercent,
        priceByExtraWeight: valorExcedente,
        maxVolume: 9999999,
        timeCost: prazo,
        country: 'BRA',
        minimumValueInsurance: 0,
      });
    });
  }

  const totalPreparedRows = rows.length - startRow;
  const totalGeneratedRows = vtexRows.length;
  const averageCost = totalGeneratedRows > 0 ? (totalCostSum / totalGeneratedRows) : 0;

  const stats: ConversionStats = {
    totalPreparedRows,
    totalGeneratedRows,
    uniqueZipsCount: uniqueZips.size,
    maxWeightKg,
    averageCost,
    timeRangeMin: minTime === Infinity ? 0 : minTime,
    timeRangeMax: maxTime === -Infinity ? 0 : maxTime,
  };

  return {
    vtexRows,
    stats,
  };
}

/**
 * Builds and downloads the standardized VTEX spreadsheet
 */
export function exportToVtexExcel(vtexRows: VtexShippingRow[]): { blob: Blob; filename: string } {
  // VTEX standard rows
  const headerVtex = [
    [null, 'Necessário preenchimento', null, null, null, null, null, null, null, null, null, null, null],
    [null, 'Não necessário o preenchimento', null, null, null, null, null, null, null, null, null, null, null],
    [null, null, null, null, null, null, null, null, null, null, null, null, null],
    [
      null,
      'RANGE CEP - INICIO',
      'RANGE CEP - FIM',
      null,
      'PESO INICIAL',
      'PESO FINAL',
      'FRETE POR PESO',
      'GRIS+ADV',
      'PESO EXCEDENTE',
      'VOLUME MAXIMO',
      'PRAZO ENTREGA',
      null,
      'MIN. VALOR POR NOTA'
    ],
    [
      null,
      'ZipCodeStart',
      'ZipCodeEnd',
      'PolygonName',
      'WeightStart',
      'WeightEnd',
      'AbsoluteMoneyCost',
      'PricePercent',
      'PriceByExtraWeight',
      'MaxVolume',
      'TimeCost',
      'Country',
      'MinimumValueInsurance'
    ]
  ];

  // Map vtex rows array to data rows
  const dataRows = vtexRows.map((row) => [
    null,
    row.zipCodeStart,
    row.zipCodeEnd,
    row.polygonName,
    row.weightStart,
    row.weightEnd,
    row.absoluteMoneyCost,
    row.pricePercent,
    row.priceByExtraWeight,
    row.maxVolume,
    row.timeCost,
    row.country,
    row.minimumValueInsurance
  ]);

  const outputRows = [...headerVtex, ...dataRows];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(outputRows);

  // Set explicit column widths for readability in Excel
  const wscols = [
    { wch: 5 },  // empty col A
    { wch: 18 }, // ZipCodeStart
    { wch: 18 }, // ZipCodeEnd
    { wch: 15 }, // PolygonName
    { wch: 12 }, // WeightStart
    { wch: 12 }, // WeightEnd
    { wch: 18 }, // AbsoluteMoneyCost
    { wch: 12 }, // PricePercent
    { wch: 18 }, // PriceByExtraWeight
    { wch: 15 }, // MaxVolume
    { wch: 15 }, // TimeCost
    { wch: 10 }, // Country
    { wch: 22 }  // MinimumValueInsurance
  ];
  ws['!cols'] = wscols;

  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

  const fileBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([fileBuffer], { type: 'application/octet-stream' });
  const filename = `planilha_vtex_fretes_${new Date().toISOString().slice(0, 10)}.xlsx`;

  return { blob, filename };
}
