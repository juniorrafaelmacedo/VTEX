/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface WeightRange {
  colIdx: number;
  label: string;
  weightKg: number;
  start: number; // in grams
  end: number;   // in grams
}

export interface SpreadsheetMapping {
  sheetName: string;
  headerRowIndex: number;
  zipCodeStartCol: number;
  zipCodeEndCol: number;
  prazoCol: number;
  excedenteCol: number;
  grisCol: number | null;
  seguroCol: number | null;
  weightCols: number[]; // Columns carrying float weights
}

export interface VtexShippingRow {
  index: number;
  zipCodeStart: string;
  zipCodeEnd: string;
  polygonName: string;
  weightStart: number;
  weightEnd: number;
  absoluteMoneyCost: number;
  pricePercent: number;
  priceByExtraWeight: number;
  maxVolume: number;
  timeCost: number;
  country: string;
  minimumValueInsurance: number;
}

export interface ConversionStats {
  totalPreparedRows: number;
  totalGeneratedRows: number;
  uniqueZipsCount: number;
  maxWeightKg: number;
  averageCost: number;
  timeRangeMin: number;
  timeRangeMax: number;
}
