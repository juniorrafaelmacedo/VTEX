/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  FileSpreadsheet,
  Download,
  RefreshCw,
  HelpCircle,
  CheckCircle2,
  AlertTriangle,
  ArrowRight,
  BookOpen,
  ChevronDown,
  ChevronUp,
  FileDown
} from 'lucide-react';

import UploadZone from './components/UploadZone';
import MappingConfigs from './components/MappingConfigs';
import StatsDashboard from './components/StatsDashboard';
import ResultsGrid from './components/ResultsGrid';
import { SpreadsheetMapping } from './types';
import {
  detectSpreadsheetDetails,
  processIntelipostToVtex,
  exportToVtexExcel,
  DEFAULT_WEIGHTS_KG
} from './utils';

export default function App() {
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [loadedFileName, setLoadedFileName] = useState<string | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [mapping, setMapping] = useState<SpreadsheetMapping | null>(null);
  const [weightsConfig, setWeightsConfig] = useState<number[]>(DEFAULT_WEIGHTS_KG);
  const [conversionError, setConversionError] = useState<string | null>(null);

  // Help Section Toggle state
  const [showGuide, setShowGuide] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  // Triggered when Excel loading completes
  const handleFileLoaded = (file: File, wb: XLSX.WorkBook) => {
    setLoadedFileName(file.name);
    setWorkbook(wb);
    setSheetNames(wb.SheetNames);
    
    // Attempt auto-detection
    const { suggestedMapping } = detectSpreadsheetDetails(wb.SheetNames, wb);
    setMapping(suggestedMapping);
    setConversionError(null);
  };

  // Perform translation in real-time when mapping configurations update
  const conversionResult = useMemo(() => {
    if (!workbook || !mapping) return null;
    try {
      const result = processIntelipostToVtex(workbook, mapping, weightsConfig);
      setConversionError(null);
      return result;
    } catch (err: any) {
      setConversionError(err?.message || 'Erro inesperado na transformação de CEP/Pesos');
      return null;
    }
  }, [workbook, mapping, weightsConfig]);

  // Clean / Reset Workspace
  const handleReset = () => {
    setWorkbook(null);
    setLoadedFileName(null);
    setSheetNames([]);
    setMapping(null);
    setWeightsConfig(DEFAULT_WEIGHTS_KG);
    setConversionError(null);
  };

  // Export spreadsheet triggered
  const handleExport = () => {
    if (!conversionResult || conversionResult.vtexRows.length === 0) return;
    setIsExporting(true);

    setTimeout(() => {
      try {
        const { blob, filename } = exportToVtexExcel(conversionResult.vtexRows);
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      } catch (err) {
        console.error('Erro na exportação', err);
      } finally {
        setIsExporting(false);
      }
    }, 600); // UI feel delay
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col font-sans">
      
      {/* Header bar */}
      <header className="bg-white border-b border-gray-150 sticky top-0 z-40 shadow-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between gap-4">
          <div className="flex items-center gap-2.5">
            <div className="p-2 rounded-xl bg-gray-900 text-white shadow-sm flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5 text-amber-500 stroke-[2]" />
            </div>
            <div>
              <h1 className="text-sm sm:text-base font-bold text-gray-900 tracking-tight leading-none mb-1">
                Conversor Intelipost para VTEX
              </h1>
              <span className="text-[10px] text-gray-400 font-medium block">
                Planilhas de frete inteligentes e sem fricção
              </span>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              id="help-guide-toggle"
              onClick={() => setShowGuide(!showGuide)}
              type="button"
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-gray-150 hover:bg-gray-50 text-xs font-semibold text-gray-700 transition cursor-pointer"
            >
              <BookOpen className="w-4 h-4 text-gray-500" />
              <span>Instruções VTEX</span>
              {showGuide ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
            </button>

            {workbook && (
              <button
                id="reset-workspace-button"
                onClick={handleReset}
                type="button"
                className="flex items-center justify-center p-1.5 rounded-lg border border-gray-200 hover:bg-rose-50 hover:text-rose-700 text-gray-500 transition cursor-pointer"
                title="Limpar Planilha e Recomeçar"
              >
                <RefreshCw className="w-4 h-4 animate-hover-spin" />
              </button>
            )}
          </div>
        </div>
      </header>

      {/* Accordion Guide Block */}
      {showGuide && (
        <section className="bg-white border-b border-gray-150 py-5 transition-all duration-300">
          <div className="max-w-4xl mx-auto px-4">
            <h2 className="text-sm font-bold text-gray-900 mb-2.5 uppercase tracking-wider flex items-center gap-1.5">
              <span>Passo a Passo para Importação na VTEX</span>
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs text-gray-650">
              <div className="bg-gray-55/10 rounded-xl p-3.5 border border-gray-100 flex flex-col gap-1.5">
                <span className="w-5 h-5 rounded-full bg-gray-950 text-white font-mono flex items-center justify-center text-[10px] font-bold">1</span>
                <span className="font-semibold text-gray-800">Converta o Arquivo</span>
                <p className="text-[11px] leading-relaxed text-gray-500">
                  Carregue a planilha da Intelipost (v2.5) e defina qual aba e linha representam os cabeçalhos. Os CEPs serão formatados automaticamente em <strong>00000-000</strong> e as faixas divididas em gramas.
                </p>
              </div>
              <div className="bg-gray-55/10 rounded-xl p-3.5 border border-gray-100 flex flex-col gap-1.5">
                <span className="w-5 h-5 rounded-full bg-gray-950 text-white font-mono flex items-center justify-center text-[10px] font-bold">2</span>
                <span className="font-semibold text-gray-800">Baixe o Gerado</span>
                <p className="text-[11px] leading-relaxed text-gray-500">
                  Faça o download da planilha VTEX que cumpre estritamente o layout padrão com as 13 colunas estruturadas, faixas de peso concatenadas e soma de Gris + Seguro no <strong>PricePercent</strong>.
                </p>
              </div>
              <div className="bg-gray-55/10 rounded-xl p-3.5 border border-gray-100 flex flex-col gap-1.5">
                <span className="w-5 h-5 rounded-full bg-gray-950 text-white font-mono flex items-center justify-center text-[10px] font-bold">3</span>
                <span className="font-semibold text-gray-800">Importe na VTEX</span>
                <p className="text-[11px] leading-relaxed text-gray-500">
                  Acesse o painel <strong>VTEX Admin</strong> &gt; <strong>Logística</strong> &gt; <strong>Valores de frete</strong>.Selecione a transportadora correspondente e envie a nova planilha convertida.
                </p>
              </div>
            </div>
          </div>
        </section>
      )}

      {/* Main Workspace Frame */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 flex flex-col gap-6">
        
        {!workbook ? (
          // Empty State: Prominent Drag Field and Quick Info Overview
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 my-auto items-stretch">
            <div className="lg:col-span-1">
              <UploadZone
                onFileLoaded={handleFileLoaded}
                isLoading={false}
                loadedFileName={null}
              />
            </div>

            <div className="lg:col-span-2 flex flex-col justify-between bg-white border border-gray-150 rounded-2xl p-6 shadow-sm">
              <div>
                <span className="text-xs font-semibold text-amber-600 bg-amber-50 px-2.5 py-1 rounded-full uppercase tracking-widest font-mono">
                  Compatibilidade Ativa
                </span>
                <h3 className="text-md sm:text-lg font-bold text-gray-950 mt-3 mb-2 leading-tight">
                  Automação Completa de Tabelas de Frete
                </h3>
                <p className="text-xs text-gray-550 mb-4 leading-relaxed">
                  Evite horas de retrabalho manual no Microsoft Excel separando grandes faixas de CEP e pesos. Nosso conversor realiza de forma instantânea a expansão cartesiana de rotas por quilo conforme exigido pelo modelo de planilha da VTEX.
                </p>

                <div className="space-y-3 pt-2 text-xs">
                  <div className="flex items-start gap-2.5 text-gray-700">
                    <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                    <div>
                      <strong>Ajuste de CEPs Automático:</strong> Limpa pontuações incorretas e padroniza para <strong>00000-000</strong> com zeros à esquerda.
                    </div>
                  </div>
                  <div className="flex items-start gap-2.5 text-gray-700">
                    <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                    <div>
                      <strong>Tratamento de Pesos:</strong> Converte quilos (KG) para gramas e cria de forma sequencial faixas de peso correspondentes em gramagem (ex: 0g a 250g, 251g a 500g...).
                    </div>
                  </div>
                  <div className="flex items-start gap-2.5 text-gray-700">
                    <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                    <div>
                      <strong>Valor do Ad Valorem:</strong> Soma de forma dinâmica as taxas financeiras de GRIS e Seguro Declarado.
                    </div>
                  </div>
                </div>
              </div>

              <div className="border-t border-gray-100 pt-4 mt-6 flex flex-col sm:flex-row items-center justify-between text-[11px] text-gray-400 gap-2">
                <span>Tecnologia em lote 100% segura que mantém as informações da sua empresa localizadas e protegidas na sua máquina.</span>
                <span className="font-mono text-gray-300">v1.2.0-secure</span>
              </div>
            </div>
          </div>
        ) : (
          // Active Processing State
          <div className="flex flex-col gap-6 animate-fade-in">
            {/* Split layout: sidebar and reviews */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch">
              {/* Configuration sidebar (4 columns layout) */}
              <div className="lg:col-span-4 h-full">
                {mapping && (
                  <MappingConfigs
                    sheetNames={sheetNames}
                    mapping={mapping}
                    onMappingChange={setMapping}
                    weightsConfig={weightsConfig}
                    onWeightsConfigChange={setWeightsConfig}
                    workbook={workbook}
                  />
                )}
              </div>

              {/* Outputs block (8 columns layout) */}
              <div className="lg:col-span-8 flex flex-col gap-5 h-full">
                
                {/* Visual Export trigger panel */}
                <div className="bg-white border border-gray-150 rounded-2xl p-5 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-4">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center shrink-0">
                      <CheckCircle2 className="w-5 h-5" />
                    </div>
                    <div>
                      <p className="text-xs font-semibold text-gray-900 leading-snug">
                        Suas rotas de frete foram calculadas com sucesso!
                      </p>
                      <p className="text-[11px] text-gray-500">
                        Pronto para download imediato. Altere as colunas à esquerda para mudar a tabela em tempo real.
                      </p>
                    </div>
                  </div>

                  <button
                    id="export-vtex-xlsx-button"
                    onClick={handleExport}
                    disabled={isExporting || (conversionResult?.vtexRows.length === 0)}
                    type="button"
                    className="w-full sm:w-auto flex items-center justify-center gap-2 px-6 py-2.5 bg-gray-950 text-white hover:bg-gray-800 disabled:opacity-50 font-semibold rounded-xl text-xs shadow-md shadow-gray-200 transition-all duration-150 hover:-translate-y-0.5 cursor-pointer"
                  >
                    {isExporting ? (
                      <div className="w-4 h-4 rounded-full border-2 border-slate-50 border-t-transparent animate-spin" />
                    ) : (
                      <FileDown className="w-4 h-4 stroke-[2]" />
                    )}
                    <span>{isExporting ? 'Processando...' : 'Baixar Planilha VTEX Gerada'}</span>
                  </button>
                </div>

                {/* Conversion Warnings/Exceptions displayed if mapping fails */}
                {conversionError && (
                  <div className="flex items-start gap-2.5 bg-rose-50 border border-rose-100 rounded-xl p-4 text-rose-850 text-xs">
                    <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                    <div>
                      <strong className="block font-bold">Erro na conversão:</strong>
                      <span className="text-[11px] leading-relaxed block mt-0.5 text-rose-700 font-mono">
                        {conversionError}
                      </span>
                    </div>
                  </div>
                )}

                {/* Stats Widget Display component */}
                <StatsDashboard stats={conversionResult?.stats || null} />

                {/* Table preview component */}
                {conversionResult && (
                  <div className="flex-1 min-h-0">
                    <ResultsGrid rows={conversionResult.vtexRows} />
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Footer bar */}
      <footer className="bg-white border-t border-gray-150 py-4 mt-auto">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center sm:flex sm:justify-between sm:items-center text-[10px] text-gray-400 gap-4">
          <span>&copy; {new Date().getFullYear()} Conversor VTEX Logistics. Operando integralmente no navegador.</span>
          <span className="font-semibold text-gray-300">Conformidade e precisão no envio de arquivos VTEX.</span>
        </div>
      </footer>
    </div>
  );
}
