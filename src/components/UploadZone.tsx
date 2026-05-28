/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, FileSpreadsheet, Sparkles, CheckCircle2, AlertTriangle } from 'lucide-react';

interface UploadZoneProps {
  onFileLoaded: (file: File, workbook: XLSX.WorkBook) => void;
  isLoading: boolean;
  loadedFileName: string | null;
}

export default function UploadZone({ onFileLoaded, isLoading, loadedFileName }: UploadZoneProps) {
  const [isDragActive, setIsDragActive] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setIsDragActive(true);
    } else if (e.type === 'dragleave') {
      setIsDragActive(false);
    }
  };

  const processFile = (file: File) => {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      setErrorMsg('Por favor, envie apenas arquivos de planilha excel (.xlsx, .xls)');
      return;
    }
    setErrorMsg(null);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        onFileLoaded(file, workbook);
      } catch (err) {
        setErrorMsg('Erro ao ler a planilha. Verifique se o arquivo não está corrompido.');
        console.error(err);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      processFile(e.target.files[0]);
    }
  };

  const triggerFileInput = () => {
    fileInputRef.current?.click();
  };

  /**
   * Generates a realistic mock Intelipost table that can be fed into the converter
   */
  const handleGenerateSample = () => {
    const headerRow = [
      '', '', '', '', '', '', '', '', '', '', '', // empty leading elements
    ];
    
    // Intelipost standard columns
    const columns = [
      'Código do Serviço', 'Nome da Transportadora', 'Modalidade', 'Serviço', 'Descrição', 'Região', 'Estado', 'Ativo',
      'CEP Inicial', 'CEP Final', 'Prazo de Entrega', '0.25', '0.5', '0.75', '1', '2', '3', '4', '5', '10', '20', '50',
      'VALOR EXCEDENTE', 'GRIS', 'SEGURO'
    ];

    // Mock data rows that simulate Brazil regional postcodes
    const mockRegions = [
      { start: '01000000', end: '05999999', prazo: 2, exc: 1.5, gris: 0.12, seguro: 0.35, basis: 14.5 }, // SP Capital
      { start: '11000000', end: '11999999', prazo: 3, exc: 1.8, gris: 0.15, seguro: 0.40, basis: 18.2 }, // SP Litoral
      { start: '20000000', end: '23799999', prazo: 4, exc: 2.1, gris: 0.20, seguro: 0.50, basis: 22.1 }, // RJ Capital
      { start: '30000000', end: '31999999', prazo: 3, exc: 1.7, gris: 0.14, seguro: 0.38, basis: 19.5 }, // BH Capital
      { start: '40000000', end: '41999999', prazo: 5, exc: 2.5, gris: 0.25, seguro: 0.60, basis: 29.8 }, // BA Salvador
      { start: '90000000', end: '91999999', prazo: 4, exc: 2.0, gris: 0.18, seguro: 0.45, basis: 24.5 }  // RS Porto Alegre
    ];

    // Prepare rows. Intelipost headers typically are on Row index 3 (index 3 is the 4th row)
    const worksheetData: any[][] = [
      ['RELATÓRIO DE TAXAS DE FRETE INTELIPOST'],
      ['Exportado em: 2026-05-28'],
      [], // Empty row
      columns // Headers on row idx 3
    ];

    mockRegions.forEach((reg, rank) => {
      // Create price points for specified weights
      const pricesByWeight = [
        reg.basis,         // 0.25 KG
        reg.basis * 1.15,  // 0.5 KG
        reg.basis * 1.30,  // 0.75 KG
        reg.basis * 1.45,  // 1.0 KG
        reg.basis * 1.80,  // 2.0 KG
        reg.basis * 2.15,  // 3.0 KG
        reg.basis * 2.40,  // 4.0 KG
        reg.basis * 2.70,  // 5.0 KG
        reg.basis * 3.80,  // 10.0 KG
        reg.basis * 5.20,  // 20.0 KG
        reg.basis * 8.50   // 50.0 KG
      ].map(p => parseFloat(p.toFixed(2)));

      worksheetData.push([
        `SERV-${100 + rank}`, 'Transportadora Alfa', 'Adicional', 'Rápido', 'Entrega Expressa Regional', 'Sudeste', 'UF', 'SIM',
        reg.start, reg.end, reg.prazo,
        ...pricesByWeight,
        reg.exc, reg.gris, reg.seguro
      ]);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(worksheetData);
    XLSX.utils.book_append_sheet(wb, ws, '2.5');

    // Create a virtual file to pass
    const fileBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const sampleBlob = new Blob([fileBuffer], { type: 'application/octet-stream' });
    const sampleFile = new File([sampleBlob], 'modelo_tabela_intelipost_exemplo.xlsx', {
      type: 'application/octet-stream'
    });

    onFileLoaded(sampleFile, wb);
  };

  return (
    <div className="bg-white rounded-2xl border border-gray-150 p-6 shadow-sm flex flex-col h-full justify-between gap-5 transition-all duration-300">
      <div>
        <div className="flex items-center gap-2 mb-1">
          <FileSpreadsheet className="w-5 h-5 text-gray-700" />
          <h2 className="text-md font-semibold text-gray-900">1. Upload da Planilha</h2>
        </div>
        <p className="text-xs text-gray-500 mb-4">
          Arraste a sua planilha Intelipost original ou use nossa planilha inteligente de testes.
        </p>
      </div>

      <div
        id="drag-drop-zone"
        onDragEnter={handleDrag}
        onDragOver={handleDrag}
        onDragLeave={handleDrag}
        onDrop={handleDrop}
        onClick={triggerFileInput}
        className={`relative flex-1 min-h-[170px] border-2 border-dashed rounded-xl flex flex-col items-center justify-center p-5 text-center cursor-pointer transition-all duration-200 ${
          isDragActive
            ? 'border-gray-800 bg-gray-50'
            : loadedFileName
            ? 'border-emerald-200 bg-emerald-50/20 hover:bg-emerald-50/35'
            : 'border-gray-200 bg-gray-50/40 hover:bg-gray-50/70 hover:border-gray-300'
        }`}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          onChange={handleChange}
          className="hidden"
          id="spreadsheet-file-input"
        />

        {isLoading ? (
          <div className="flex flex-col items-center gap-2">
            <div className="w-8 h-8 rounded-full border-2 border-slate-200 border-t-slate-800 animate-spin" />
            <p className="font-medium text-xs text-gray-700">Lendo arquivos e calculando colunas...</p>
          </div>
        ) : loadedFileName ? (
          <div className="flex flex-col items-center gap-2">
            <div className="w-10 h-10 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-600 shadow-sm animate-pulse">
              <CheckCircle2 className="w-5 h-5" />
            </div>
            <p className="font-semibold text-xs text-emerald-900 truncate max-w-[280px]">
              {loadedFileName}
            </p>
            <span className="text-[10px] text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded-full font-medium">
              Carregado com sucesso
            </span>
          </div>
        ) : (
          <div className="flex flex-col items-center">
            <div className="w-10 h-10 rounded-full bg-gray-100 flex items-center justify-center mb-3 text-gray-500">
              <UploadCloud className="w-5 h-5" />
            </div>
            <p className="font-medium text-xs text-gray-800 mb-1">
              Arraste a planilha Intelipost aqui
            </p>
            <p className="text-[10px] text-gray-400">
              Formatos aceitos: .xlsx ou .xls
            </p>
            <button
              type="button"
              id="select-file-button"
              className="mt-4 px-3 py-1.5 text-xs font-semibold text-white bg-gray-900 rounded-lg hover:bg-gray-800 transition duration-150 shadow-sm cursor-pointer"
            >
              Selecionar arquivo
            </button>
          </div>
        )}
      </div>

      {errorMsg && (
        <div className="flex items-start gap-2 bg-rose-50 border border-rose-100 rounded-lg p-3 text-rose-800 text-xs">
          <AlertTriangle className="w-4 h-4 text-rose-500 shrink-0 mt-0.5" />
          <span>{errorMsg}</span>
        </div>
      )}

      <div className="border-t border-gray-100 pt-3">
        <button
          onClick={handleGenerateSample}
          type="button"
          id="generate-sample-button"
          className="w-full flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-medium text-gray-700 bg-amber-55/15 hover:bg-amber-60/25 border border-amber-200/40 rounded-lg transition duration-150 cursor-pointer"
        >
          <Sparkles className="w-3.5 h-3.5 text-amber-600" />
          <span>Usar planilha de simulação (Completa)</span>
        </button>
      </div>
    </div>
  );
}
