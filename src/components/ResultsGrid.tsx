/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import { VtexShippingRow } from '../types';
import { Search, ChevronLeft, ChevronRight, Table, HelpCircle } from 'lucide-react';

interface ResultsGridProps {
  rows: VtexShippingRow[];
}

export default function ResultsGrid({ rows }: ResultsGridProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);

  const formatPrice = (val: number) => {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);
  };

  // Filter rows list based on query (by Cep start or end)
  const filteredRows = useMemo(() => {
    setCurrentPage(1); // Reset page on new filter
    if (!searchTerm.trim()) return rows;
    const cleanSearch = searchTerm.replace(/\D/g, '');
    return rows.filter((row) => {
      const cleanStart = row.zipCodeStart.replace(/\D/g, '');
      const cleanEnd = row.zipCodeEnd.replace(/\D/g, '');
      return cleanStart.includes(cleanSearch) || cleanEnd.includes(cleanSearch);
    });
  }, [rows, searchTerm]);

  // Paginated Rows
  const totalPages = Math.ceil(filteredRows.length / pageSize) || 1;
  const paginatedRows = useMemo(() => {
    const startIdx = (currentPage - 1) * pageSize;
    return filteredRows.slice(startIdx, startIdx + pageSize);
  }, [filteredRows, currentPage, pageSize]);

  return (
    <div className="bg-white rounded-2xl border border-gray-150 shadow-sm overflow-hidden flex flex-col h-full min-h-[450px]">
      {/* Table Header Filter Action Bar */}
      <div className="p-4 border-b border-gray-150 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 bg-gray-50/50">
        <div className="flex items-center gap-2">
          <Table className="w-4 h-4 text-gray-700 font-bold" />
          <h3 className="text-sm font-semibold text-gray-900">Pré-visualização da Planilha VTEX</h3>
          <span className="text-[10px] bg-slate-100 text-slate-800 font-semibold px-2 py-0.5 rounded-full font-mono">
            {filteredRows.length.toLocaleString('pt-BR')} de {rows.length.toLocaleString('pt-BR')} registros
          </span>
        </div>

        {/* Searching input */}
        <div className="relative max-w-xs w-full">
          <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            id="vtex-cep-search"
            type="text"
            placeholder="Filtrar por CEP (inicial ou fim)..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full text-xs pl-9 pr-3 py-2 bg-white border border-gray-200 rounded-lg outline-none focus:border-gray-800 transition shadow-inner"
          />
        </div>
      </div>

      {/* Main Table view */}
      <div className="flex-1 overflow-x-auto">
        <table className="w-full border-collapse text-left">
          <thead>
            <tr className="border-b border-gray-150 bg-gray-50 text-[10px] font-bold text-gray-500 uppercase tracking-wider font-sans">
              <th className="py-3 px-4 w-12 text-center">Nº</th>
              <th className="py-3 px-4">CEP Inicial</th>
              <th className="py-3 px-4">CEP Final</th>
              <th className="py-3 px-4 text-right">Peso Inicial (g)</th>
              <th className="py-3 px-4 text-right">Peso Final (g)</th>
              <th className="py-3 px-4 text-right">Custo Fixo</th>
              <th className="py-3 px-4 text-right">Gris+Adv (%)</th>
              <th className="py-3 px-4 text-right">Kg Excedente</th>
              <th className="py-3 px-4 text-center">Prazo (Dias)</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100 text-xs text-gray-700">
            {paginatedRows.map((row) => (
              <tr key={row.index} className="hover:bg-gray-50/40 transition">
                <td className="py-2.5 px-4 text-center font-mono text-[10px] text-gray-400">
                  {row.index}
                </td>
                <td className="py-2.5 px-4 font-mono font-medium text-gray-950">
                  {row.zipCodeStart}
                </td>
                <td className="py-2.5 px-4 font-mono font-medium text-gray-950">
                  {row.zipCodeEnd}
                </td>
                <td className="py-2.5 px-4 text-right font-mono">
                  {row.weightStart.toLocaleString('pt-BR')}
                </td>
                <td className="py-2.5 px-4 text-right font-mono">
                  {row.weightEnd.toLocaleString('pt-BR')}
                </td>
                <td className="py-2.5 px-4 text-right font-mono font-semibold text-emerald-800">
                  {formatPrice(row.absoluteMoneyCost)}
                </td>
                <td className="py-2.5 px-4 text-right font-mono text-gray-600">
                  {row.pricePercent.toFixed(2)}%
                </td>
                <td className="py-2.5 px-4 text-right font-mono text-gray-600">
                  {formatPrice(row.priceByExtraWeight)}
                </td>
                <td className="py-2.5 px-4 text-center font-medium">
                  <span className="inline-flex items-center justify-center bg-indigo-50 text-indigo-800 px-2 py-0.5 rounded font-mono text-[11px] font-semibold">
                    {row.timeCost} d
                  </span>
                </td>
              </tr>
            ))}

            {paginatedRows.length === 0 && (
              <tr>
                <td colSpan={9} className="py-12 text-center text-gray-400">
                  <div className="flex flex-col items-center gap-1">
                    <HelpCircle className="w-6 h-6 opacity-30" />
                    <p className="text-xs">Nenhum registro encontrado para "{searchTerm}"</p>
                  </div>
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination Actions */}
      <div className="p-3 border-t border-gray-150 flex flex-col sm:flex-row items-center justify-between gap-3 bg-gray-50/50">
        <div className="flex items-center gap-1.5 text-xs text-gray-500">
          <label htmlFor="items-per-page-select" className="font-medium">Itens por página:</label>
          <select
            id="items-per-page-select"
            value={pageSize}
            onChange={(e) => {
              setPageSize(Number(e.target.value));
              setCurrentPage(1);
            }}
            className="bg-white border border-gray-200 rounded px-2 py-1 text-xs cursor-pointer focus:border-gray-800"
          >
            {[10, 15, 25, 50, 100].map(size => (
              <option key={size} value={size}>{size}</option>
            ))}
          </select>
        </div>

        <div className="flex items-center gap-3">
          <span className="text-xs text-gray-600 font-mono">
            Pág. <strong>{currentPage}</strong> de {totalPages}
          </span>
          <div className="flex items-center gap-1">
            <button
              id="pagination-prev"
              type="button"
              disabled={currentPage === 1}
              onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
              className="p-1.5 rounded-lg border border-gray-200 bg-white hover:bg-gray-50 text-gray-700 disabled:opacity-40 disabled:hover:bg-white transition cursor-pointer"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
            <button
              id="pagination-next"
              type="button"
              disabled={currentPage === totalPages}
              onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
              className="p-1.5 rounded-lg border border-gray-200 bg-white hover:bg-gray-50 text-gray-700 disabled:opacity-40 disabled:hover:bg-white transition cursor-pointer"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
