/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { ConversionStats } from '../types';
import { Database, TrendingUp, Compass, Columns, Clock, DollarSign } from 'lucide-react';

interface StatsDashboardProps {
  stats: ConversionStats | null;
}

export default function StatsDashboard({ stats }: StatsDashboardProps) {
  if (!stats) {
    return (
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4">
        {[...Array(6)].map((_, i) => (
          <div key={i} className="bg-white rounded-xl border border-gray-150 p-4 h-20 animate-pulse flex flex-col justify-between">
            <div className="h-2 w-10 bg-gray-100 rounded" />
            <div className="h-4 w-16 bg-gray-150 rounded" />
          </div>
        ))}
      </div>
    );
  }

  const formatPrice = (val: number) => {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);
  };

  const statItems = [
    {
      id: "stat-total-intelipost",
      label: 'Linhas Intelipost',
      value: stats.totalPreparedRows.toLocaleString('pt-BR'),
      desc: 'Rotas de origem lidas',
      icon: Database,
      iconColor: 'text-amber-600',
      bgColor: 'bg-amber-50/50',
    },
    {
      id: "stat-total-vtex",
      label: 'Linhas VTEX Geradas',
      value: stats.totalGeneratedRows.toLocaleString('pt-BR'),
      desc: 'Rotas x faixas de peso',
      icon: TrendingUp,
      iconColor: 'text-emerald-600',
      bgColor: 'bg-emerald-50/50',
    },
    {
      id: "stat-unique-ceps",
      label: 'Áreas de CEP',
      value: stats.uniqueZipsCount.toLocaleString('pt-BR'),
      desc: 'CEPs distintos mapeados',
      icon: Compass,
      iconColor: 'text-sky-600',
      bgColor: 'bg-sky-50/50',
    },
    {
      id: "stat-avg-cost",
      label: 'Custo Médio',
      value: formatPrice(stats.averageCost),
      desc: 'Média de frete gerada',
      icon: DollarSign,
      iconColor: 'text-purple-600',
      bgColor: 'bg-purple-50/50',
    },
    {
      id: "stat-max-weight",
      label: 'Limite de Peso',
      value: `${stats.maxWeightKg} KG`,
      desc: 'Maior peso encontrado',
      icon: Columns,
      iconColor: 'text-orange-600',
      bgColor: 'bg-orange-50/50',
    },
    {
      id: "stat-transit-times",
      label: 'Prazo Mín/Máx',
      value: `${stats.timeRangeMin} - ${stats.timeRangeMax} d`,
      desc: 'Dias úteis declarados',
      icon: Clock,
      iconColor: 'text-indigo-600',
      bgColor: 'bg-indigo-50/50',
    },
  ];

  return (
    <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3.5">
      {statItems.map((item) => {
        const Icon = item.icon;
        return (
          <div
            key={item.label}
            id={item.id}
            className="bg-white rounded-xl border border-gray-150 p-4 shadow-sm flex flex-col justify-between hover:shadow-md transition-shadow duration-200"
          >
            <div className="flex items-center justify-between gap-1">
              <span className="text-[10px] font-bold text-gray-400 uppercase tracking-wider truncate">
                {item.label}
              </span>
              <div className={`p-1 rounded ${item.bgColor} ${item.iconColor} shrink-0`}>
                <Icon className="w-3.5 h-3.5 stroke-[2.5]" />
              </div>
            </div>
            <div className="mt-2.5">
              <span className="text-lg font-bold text-gray-900 font-mono tracking-tight block">
                {item.value}
              </span>
              <span className="text-[9px] text-gray-400 block truncate">
                {item.desc}
              </span>
            </div>
          </div>
        );
      })}
    </div>
  );
}
