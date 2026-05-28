/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Sliders, HelpCircle, AlertCircle, Plus, Trash2, List, Grid, ChevronDown, ChevronUp } from 'lucide-react';
import { SpreadsheetMapping } from '../types';

interface MappingConfigsProps {
  sheetNames: string[];
  mapping: SpreadsheetMapping;
  onMappingChange: (m: SpreadsheetMapping) => void;
  weightsConfig: number[];
  onWeightsConfigChange: (w: number[]) => void;
  workbook: XLSX.WorkBook | null;
}

export default function MappingConfigs({
  sheetNames,
  mapping,
  onMappingChange,
  weightsConfig,
  onWeightsConfigChange,
  workbook
}: MappingConfigsProps) {
  const [showRowInspector, setShowRowInspector] = useState(false);
  const [newWeightStr, setNewWeightStr] = useState('');

  if (!workbook) {
    return (
      <div className="bg-white rounded-2xl border border-gray-150 p-6 shadow-sm h-full flex flex-col justify-center items-center text-center p-8 text-gray-400">
        <Sliders className="w-8 h-8 opacity-30 mb-2" />
        <p className="text-xs">Aguardando envio da planilha Intelipost para configurar colunas.</p>
      </div>
    );
  }

  // Get first 10 rows for inspection and mapping preview
  const sheet = workbook.Sheets[mapping.sheetName];
  const allRows = sheet ? XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1 }) : [];
  const previewRows = allRows.slice(0, 10);
  const currentHeaders = allRows[mapping.headerRowIndex] || [];

  const handleSheetChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const nextSheet = e.target.value;
    // Attempt auto-detect on the new sheet
    const nextSheetObj = workbook.Sheets[nextSheet];
    const nextRows = XLSX.utils.sheet_to_json<any[]>(nextSheetObj, { header: 1 });
    
    // Guess best row
    let bestRow = 0;
    let maxCols = 0;
    for (let i = 0; i < Math.min(nextRows.length, 10); i++) {
      if (nextRows[i] && nextRows[i].length > maxCols) {
        maxCols = nextRows[i].length;
        bestRow = i;
      }
    }

    onMappingChange({
      ...mapping,
      sheetName: nextSheet,
      headerRowIndex: bestRow,
    });
  };

  const selectHeaderRow = (rowIndex: number) => {
    // When changing header row, recalculate auto-mappings based on the new headers
    const newHeaders = allRows[rowIndex] || [];
    
    let zipCodeStartCol = mapping.zipCodeStartCol;
    let zipCodeEndCol = mapping.zipCodeEndCol;
    let prazoCol = mapping.prazoCol;
    let excedenteCol = mapping.excedenteCol;
    let grisCol = mapping.grisCol;
    let seguroCol = mapping.seguroCol;
    const weightCols: number[] = [];

    newHeaders.forEach((cell, idx) => {
      const valStr = String(cell || '').toLowerCase().trim();
      if (valStr.includes('cep') && (valStr.includes('inicial') || valStr.includes('inicio') || valStr.includes('início') || valStr.includes('start'))) {
        zipCodeStartCol = idx;
      } else if (valStr.includes('cep') && (valStr.includes('final') || valStr.includes('fim') || valStr.includes('end'))) {
        zipCodeEndCol = idx;
      } else if (valStr === 'prazo' || valStr.includes('prazo de entrega') || valStr.includes('timecost')) {
        prazoCol = idx;
      } else if (valStr.includes('excedente') || valStr.includes('extra') || valStr.includes('adicional')) {
        excedenteCol = idx;
      } else if (valStr.includes('gris')) {
        grisCol = idx;
      } else if (valStr.includes('seguro') || valStr.includes('ad valorem') || valStr.includes('advalorem')) {
        seguroCol = idx;
      }

      // Weights detector
      const parsed = parseFloat(valStr.replace(',', '.'));
      if (!isNaN(parsed) && parsed > 0 && parsed <= 500) {
        weightCols.push(idx);
      }
    });

    onMappingChange({
      ...mapping,
      headerRowIndex: rowIndex,
      zipCodeStartCol,
      zipCodeEndCol,
      prazoCol,
      excedenteCol,
      grisCol,
      seguroCol,
      weightCols,
    });
  };

  const handleColMapperChange = (field: keyof SpreadsheetMapping, value: string) => {
    let parsedVal: any = value === 'none' ? null : parseInt(value, 10);
    if (isNaN(parsedVal)) parsedVal = null;
    onMappingChange({
      ...mapping,
      [field]: parsedVal
    });
  };

  const handleAddWeight = (e: React.FormEvent) => {
    e.preventDefault();
    const weightFloat = parseFloat(newWeightStr.replace(',', '.'));
    if (!isNaN(weightFloat) && weightFloat > 0 && !weightsConfig.includes(weightFloat)) {
      const updated = [...weightsConfig, weightFloat].sort((a, b) => a - b);
      onWeightsConfigChange(updated);
      setNewWeightStr('');
    }
  };

  const handleRemoveWeight = (weight: number) => {
    onWeightsConfigChange(weightsConfig.filter(w => w !== weight));
  };

  return (
    <div className="bg-white rounded-2xl border border-gray-150 p-6 shadow-sm flex flex-col gap-5 h-full overflow-y-auto">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-2">
          <Sliders className="w-5 h-5 text-gray-700" />
          <h2 className="text-md font-semibold text-gray-900">2. Parâmetros & Mapeamento</h2>
        </div>
        <span className="text-[10px] uppercase font-bold text-gray-400 font-mono tracking-wider bg-gray-50 px-2 py-0.5 rounded">
          Intelipost v2.5
        </span>
      </div>

      {/* Sheet Picker */}
      <div className="flex flex-col gap-1.5">
        <label htmlFor="sheet-select-dropdown" className="text-xs font-semibold text-gray-700 flex items-center justify-between">
          <span>Aba da Planilha</span>
          <span className="text-[10px] text-gray-400 font-normal">Excel Sheets</span>
        </label>
        <select
          id="sheet-select-dropdown"
          value={mapping.sheetName}
          onChange={handleSheetChange}
          className="w-full text-xs bg-gray-50 border border-gray-200 rounded-lg p-2.5 outline-none focus:border-gray-800 transition cursor-pointer"
        >
          {sheetNames.map((name) => (
            <option key={name} value={name}>
              {name}
            </option>
          ))}
        </select>
      </div>

      {/* Accordion Inspector of spreadsheet header row */}
      <div className="border border-gray-100 rounded-xl overflow-hidden bg-gray-50/20">
        <button
          onClick={() => setShowRowInspector(!showRowInspector)}
          type="button"
          id="accordion-inspect-rows"
          className="w-full flex items-center justify-between p-3 text-xs font-semibold text-gray-700 hover:bg-gray-50 transition cursor-pointer"
        >
          <div className="flex items-center gap-1.5">
            <Grid className="w-3.5 h-3.5 text-gray-500" />
            <span>Alinhamento da Linha de Cabeçalho</span>
          </div>
          <div className="flex items-center gap-2">
            <span className="text-[10px] font-mono bg-gray-100 text-gray-650 px-2 py-0.5 rounded font-bold">
              Linha Atual: {mapping.headerRowIndex + 1}
            </span>
            {showRowInspector ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </div>
        </button>

        {showRowInspector && (
          <div className="p-3 border-t border-gray-100 bg-white flex flex-col gap-2 max-h-[220px] overflow-y-auto">
            <p className="text-[10px] text-gray-400 mb-1">
              Selecione a linha que contém os títulos das colunas (ex: CEP Inicial, Prazo, Pesos, etc.):
            </p>
            {previewRows.map((row, rIdx) => {
              const isSelected = rIdx === mapping.headerRowIndex;
              return (
                <div
                  key={rIdx}
                  onClick={() => selectHeaderRow(rIdx)}
                  id={`inspect-row-${rIdx}`}
                  className={`group flex items-start justify-between gap-3 p-2 rounded-lg text-left cursor-pointer transition select-none ${
                    isSelected
                      ? 'bg-gray-900 text-white'
                      : 'hover:bg-gray-50 text-gray-700'
                  }`}
                >
                  <div className="flex-1 truncate">
                    <span className={`text-[9px] font-bold font-mono mr-1.5 ${isSelected ? 'text-gray-300' : 'text-gray-400'}`}>
                      LINHA {rIdx + 1}:
                    </span>
                    <span className="text-[10px] font-mono">
                      {Array.isArray(row)
                        ? row.slice(0, 6).map(cell => String(cell || '-')).join(' | ') + (row.length > 6 ? ' ...' : '')
                        : 'Vazia'}
                    </span>
                  </div>
                  <button
                    type="button"
                    className={`text-[9px] px-1.5 py-0.5 font-bold rounded uppercase shrink-0 transition ${
                      isSelected
                        ? 'bg-white/20 text-white hover:bg-white/30'
                        : 'bg-gray-100 text-gray-700 group-hover:bg-gray-900 group-hover:text-white'
                    }`}
                  >
                    Usar
                  </button>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* Main mapping grids */}
      <div className="flex flex-col gap-3.5 border-t border-gray-100 pt-3.5">
        <h3 className="text-xs font-bold text-gray-700 uppercase tracking-widest leading-none">
          Mapeamento de Colunas
        </h3>

        <div className="grid grid-cols-2 gap-3.5">
          {/* CEP Inicial */}
          <div className="flex flex-col gap-1">
            <label htmlFor="map-zip-start" className="text-[11px] font-semibold text-gray-600">
              CEP Inicial
            </label>
            <select
              id="map-zip-start"
              value={mapping.zipCodeStartCol}
              onChange={(e) => handleColMapperChange('zipCodeStartCol', e.target.value)}
              className="w-full text-xs bg-gray-50 border border-gray-200 rounded-lg p-2 outline-none focus:border-gray-800 transition cursor-pointer"
            >
              {currentHeaders.map((head, i) => (
                <option key={i} value={i}>
                  Col {i + 1}: {String(head || '').slice(0, 20) || `(Coluna ${i + 1})`}
                </option>
              ))}
            </select>
          </div>

          {/* CEP Final */}
          <div className="flex flex-col gap-1">
            <label htmlFor="map-zip-end" className="text-[11px] font-semibold text-gray-600">
              CEP Final
            </label>
            <select
              id="map-zip-end"
              value={mapping.zipCodeEndCol}
              onChange={(e) => handleColMapperChange('zipCodeEndCol', e.target.value)}
              className="w-full text-xs bg-gray-50 border border-gray-200 rounded-lg p-2 outline-none focus:border-gray-800 transition cursor-pointer"
            >
              {currentHeaders.map((head, i) => (
                <option key={i} value={i}>
                  Col {i + 1}: {String(head || '').slice(0, 20) || `(Coluna ${i + 1})`}
                </option>
              ))}
            </select>
          </div>

          {/* Prazo */}
          <div className="flex flex-col gap-1">
            <label htmlFor="map-prazo" className="text-[11px] font-semibold text-gray-600">
              Prazo de Entrega
            </label>
            <select
              id="map-prazo"
              value={mapping.prazoCol}
              onChange={(e) => handleColMapperChange('prazoCol', e.target.value)}
              className="w-full text-xs bg-gray-50 border border-gray-200 rounded-lg p-2 outline-none focus:border-gray-800 transition cursor-pointer"
            >
              {currentHeaders.map((head, i) => (
                <option key={i} value={i}>
                  Col {i + 1}: {String(head || '').slice(0, 20) || `(Coluna ${i + 1})`}
                </option>
              ))}
            </select>
          </div>

          {/* Valor Excedente */}
          <div className="flex flex-col gap-1">
            <label htmlFor="map-excedente" className="text-[11px] font-semibold text-gray-600">
              Valor Excedente (R$/KG)
            </label>
            <select
              id="map-excedente"
              value={mapping.excedenteCol}
              onChange={(e) => handleColMapperChange('excedenteCol', e.target.value)}
              className="w-full text-xs bg-gray-50 border border-gray-200 rounded-lg p-2 outline-none focus:border-gray-800 transition cursor-pointer"
            >
              {currentHeaders.map((head, i) => (
                <option key={i} value={i}>
                  Col {i + 1}: {String(head || '').slice(0, 20) || `(Coluna ${i + 1})`}
                </option>
              ))}
            </select>
          </div>

          {/* GRIS */}
          <div className="flex flex-col gap-1">
            <label htmlFor="map-gris" className="text-[11px] font-semibold text-gray-600">
              Taxa GRIS (%)
            </label>
            <select
              id="map-gris"
              value={mapping.grisCol ?? 'none'}
              onChange={(e) => handleColMapperChange('grisCol', e.target.value)}
              className="w-full text-xs bg-gray-50 border border-gray-200 rounded-lg p-2 outline-none focus:border-gray-800 transition cursor-pointer"
            >
              <option value="none">Nenhum (Valor = 0%)</option>
              {currentHeaders.map((head, i) => (
                <option key={i} value={i}>
                  Col {i + 1}: {String(head || '').slice(0, 20) || `(Coluna ${i + 1})`}
                </option>
              ))}
            </select>
          </div>

          {/* Seguro */}
          <div className="flex flex-col gap-1">
            <label htmlFor="map-seguro" className="text-[11px] font-semibold text-gray-600">
              Seguro / Ad Valorem (%)
            </label>
            <select
              id="map-seguro"
              value={mapping.seguroCol ?? 'none'}
              onChange={(e) => handleColMapperChange('seguroCol', e.target.value)}
              className="w-full text-xs bg-gray-50 border border-gray-200 rounded-lg p-2 outline-none focus:border-gray-800 transition cursor-pointer"
            >
              <option value="none">Nenhum (Valor = 0%)</option>
              {currentHeaders.map((head, i) => (
                <option key={i} value={i}>
                  Col {i + 1}: {String(head || '').slice(0, 20) || `(Coluna ${i + 1})`}
                </option>
              ))}
            </select>
          </div>
        </div>

        <div className="flex items-start gap-1.5 bg-gray-50 border border-gray-100 rounded-lg p-2.5 text-[10px] text-gray-500">
          <AlertCircle className="w-3.5 h-3.5 text-gray-400 shrink-0 mt-0.5" />
          <span>
            A VTEX utiliza <strong>PricePercent</strong> somando GRIS + Seguro, que incide proporcionalmente sobre o frete. Esses valores serão somados em tempo real na coluna <strong>GRIS+ADV</strong>.
          </span>
        </div>
      </div>

      {/* Weights Config */}
      <div className="flex flex-col gap-2.5 border-t border-gray-100 pt-3.5">
        <div className="flex items-center justify-between">
          <label className="text-xs font-bold text-gray-700 uppercase tracking-widest leading-none">
            Faixas de Peso Extrair (KG)
          </label>
          <span className="text-[10px] font-mono font-semibold text-gray-400">
            {weightsConfig.length} faixas ativas
          </span>
        </div>

        <form onSubmit={handleAddWeight} className="flex gap-2">
          <input
            id="weight-input-field"
            type="text"
            placeholder="Ex: 12.5"
            value={newWeightStr}
            onChange={(e) => setNewWeightStr(e.target.value)}
            className="flex-1 text-xs bg-gray-50 border border-gray-200 rounded-lg px-2.5 py-1.5 outline-none focus:border-gray-800 transition"
          />
          <button
            type="submit"
            id="add-weight-button"
            className="px-3 bg-gray-950 text-white rounded-lg hover:bg-gray-800 transition text-xs font-semibold flex items-center justify-center gap-1 cursor-pointer"
          >
            <Plus className="w-4 h-4" />
            <span>Inserir</span>
          </button>
        </form>

        <div className="flex flex-wrap gap-1.5 max-h-[120px] overflow-y-auto p-1 border border-dashed border-gray-150 rounded-lg bg-gray-50/20">
          {weightsConfig.map((w) => (
            <div
              key={w}
              className="flex items-center gap-1 bg-gray-100 border border-gray-200 rounded px-2 py-0.5 text-[10px] font-mono font-medium text-gray-700 hover:border-gray-300 transition"
            >
              <span>{w} KG</span>
              <button
                type="button"
                id={`remove-weight-${w}`}
                onClick={() => handleRemoveWeight(w)}
                className="hover:text-rose-600 transition duration-100 focus:outline-none cursor-pointer"
              >
                <Trash2 className="w-3 h-3 ml-0.5 stroke-[2]" />
              </button>
            </div>
          ))}
          {weightsConfig.length === 0 && (
            <span className="text-[10px] text-gray-400 p-1">Nenhum peso adicionado. O conversor extrairá todas as colunas numéricas candidatas.</span>
          )}
        </div>
      </div>
    </div>
  );
}
