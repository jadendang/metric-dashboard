import React, { useState } from 'react';
import { BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from 'recharts';
import { TrendingUp, TrendingDown, AlertCircle, Upload } from 'lucide-react';
import * as XLSX from 'xlsx';
import './index.css';


const Dashboard = () => {
  const [data, setData] = useState(null);
  const [rankingMetric, setRankingMetric] = useState('revenue2024');
  const [topN, setTopN] = useState(10);
  const [error, setError] = useState(null);
  const [fileName, setFileName] = useState(null);

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setError(null);
      setFileName(file.name);
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      
      const rawArray = [];
      for (let R = range.s.r; R <= range.e.r; ++R) {
        const row = [];
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cell = worksheet[XLSX.utils.encode_cell({ r: R, c: C })];
          row.push(cell ? cell.v : undefined);
        }
        rawArray.push(row);
      }

      const manufacturers = [];
      for (let i = 2; i < rawArray.length; i++) {
        const row = rawArray[i];
        const name = String(row[0] || '').trim();
        
        if (!name || name.length < 2) continue;

        manufacturers.push({
          manufacturer: name,
          revenue2023: parseFloat(row[1]) || 0,
          cost2023: parseFloat(row[2]) || 0,
          gp2023: parseFloat(row[3]) || 0,
          revenue2024: parseFloat(row[5]) || 0,
          cost2024: parseFloat(row[6]) || 0,
          gp2024: parseFloat(row[7]) || 0,
          revenue2025: parseFloat(row[9]) || 0,
          cost2025: parseFloat(row[10]) || 0,
          gp2025: parseFloat(row[11]) || 0,
        });
      }

      console.log(`✓ Loaded ${manufacturers.length} manufacturers`);
      console.log('First:', manufacturers[0]);
      setData(manufacturers);
    } catch (err) {
      setError(`Error: ${err.message}`);
      console.error(err);
    }
  };

  const getTopManufacturers = () => {
    if (!data) return [];

    let sorted = [...data];
    if (rankingMetric === 'revenue2024') {
      sorted.sort((a, b) => b.revenue2024 - a.revenue2024);
    } else if (rankingMetric === 'gp2024') {
      sorted.sort((a, b) => b.gp2024 - a.gp2024);
    } else if (rankingMetric === 'gpPct2024') {
      sorted.sort((a, b) => {
        const bPct = b.revenue2024 > 0 ? (b.gp2024 / b.revenue2024) * 100 : 0;
        const aPct = a.revenue2024 > 0 ? (a.gp2024 / a.revenue2024) * 100 : 0;
        return bPct - aPct;
      });
    }

    return sorted.slice(0, topN).map((item, idx) => ({
      ...item,
      rank: idx + 1,
      change2324: item.revenue2024 - item.revenue2023,
      changePct2324: item.revenue2023 > 0 ? ((item.revenue2024 - item.revenue2023) / item.revenue2023) * 100 : 0,
      change2425: item.revenue2025 - item.revenue2024,
      changePct2425: item.revenue2024 > 0 ? ((item.revenue2025 - item.revenue2024) / item.revenue2024) * 100 : 0,
      gpPct2023: item.revenue2023 > 0 ? (item.gp2023 / item.revenue2023) * 100 : 0,
      gpPct2024: item.revenue2024 > 0 ? (item.gp2024 / item.revenue2024) * 100 : 0,
      gpPct2025: item.revenue2025 > 0 ? (item.gp2025 / item.revenue2025) * 100 : 0,
    }));
  };

  const formatCurrency = (value) => {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    }).format(value);
  };

  const topMfg = getTopManufacturers();
  const chartData = topMfg.map((item) => ({
    name: item.manufacturer.substring(0, 15),
    '2023': item.revenue2023,
    '2024': item.revenue2024,
    '2025': item.revenue2025,
  }));

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-6">
      <div className="max-w-7xl mx-auto">
        <h1 className="text-4xl font-bold text-gray-900 mb-2">Sales Dashboard</h1>
        <p className="text-gray-600 mb-8">2023, 2024, and 2025 Revenue Comparison</p>

        <div className="bg-white rounded-xl shadow-lg p-6 mb-8 border-2 border-dashed border-blue-300">
          <label className="flex items-center gap-2 cursor-pointer bg-blue-600 hover:bg-blue-700 text-white px-6 py-3 rounded-lg font-medium transition-colors w-fit">
            <Upload className="h-5 w-5" />
            Upload Excel File
            <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="hidden" />
          </label>
          {fileName && <span className="text-gray-600 ml-4">{fileName} ✓</span>}
        </div>

        {error && (
          <div className="bg-red-50 border-l-4 border-red-500 p-4 mb-6 rounded-lg">
            <p className="text-red-800">{error}</p>
          </div>
        )}

        {data && (
          <>
            <div className="bg-white rounded-xl shadow-lg p-6 mb-8">
              <div className="flex flex-wrap gap-6">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Rank By</label>
                  <select value={rankingMetric} onChange={(e) => setRankingMetric(e.target.value)} className="bg-gray-50 border border-gray-300 rounded-lg px-4 py-2">
                    <option value="revenue2024">Revenue 2024</option>
                    <option value="gp2024">Gross Profit $</option>
                    <option value="gpPct2024">Gross Profit %</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Show Top</label>
                  <select value={topN} onChange={(e) => setTopN(parseInt(e.target.value))} className="bg-gray-50 border border-gray-300 rounded-lg px-4 py-2">
                    <option value={5}>5</option>
                    <option value={10}>10</option>
                    <option value={15}>15</option>
                  </select>
                </div>
                <div className="flex items-end">
                  <span className="text-gray-600">Total: <strong>{data.length}</strong> manufacturers</span>
                </div>
              </div>
            </div>

            <div className="bg-white rounded-xl shadow-lg p-6 mb-8">
              <h2 className="text-2xl font-bold text-gray-900 mb-6">Revenue by Year</h2>
              <ResponsiveContainer width="100%" height={400}>
                <BarChart data={chartData} margin={{ top: 20, right: 30, left: 20, bottom: 100 }}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="name" angle={-45} textAnchor="end" height={120} tick={{ fontSize: 12 }} />
                  <YAxis domain={[0, 55000000]} tickFormatter={(v) => `${(v / 1000000).toFixed(0)}M`} />
                  <Tooltip formatter={(v) => formatCurrency(v)} />
                  <Legend />
                  <Bar dataKey="2023" fill="#D1D5DB" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="2024" fill="#3B82F6" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="2025" fill="#10B981" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8">
              <div className="bg-white rounded-xl shadow-lg p-6">
                <h2 className="text-2xl font-bold text-gray-900 mb-6">2024 Revenue Distribution</h2>
                <ResponsiveContainer width="100%" height={300}>
                  <PieChart>
                    <Pie data={topMfg} dataKey="revenue2024" nameKey="manufacturer" cx="50%" cy="50%" outerRadius={100} label={({ manufacturer, revenue2024 }) => `${manufacturer.substring(0, 12)}: ${formatCurrency(revenue2024)}`}>
                      {topMfg.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={['#3B82F6', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899', '#06B6D4', '#14B8A6', '#F97316', '#6366F1'][index % 10]} />
                      ))}
                    </Pie>
                    <Tooltip formatter={(value) => formatCurrency(value)} />
                  </PieChart>
                </ResponsiveContainer>
              </div>

              <div className="bg-white rounded-xl shadow-lg p-6">
                <h2 className="text-2xl font-bold text-gray-900 mb-6">2025 Revenue Distribution</h2>
                <ResponsiveContainer width="100%" height={300}>
                  <PieChart>
                    <Pie data={topMfg} dataKey="revenue2025" nameKey="manufacturer" cx="50%" cy="50%" outerRadius={100} label={({ manufacturer, revenue2025 }) => `${manufacturer.substring(0, 12)}: ${formatCurrency(revenue2025)}`}>
                      {topMfg.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={['#3B82F6', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899', '#06B6D4', '#14B8A6', '#F97316', '#6366F1'][index % 10]} />
                      ))}
                    </Pie>
                    <Tooltip formatter={(value) => formatCurrency(value)} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="bg-white rounded-xl shadow-lg overflow-hidden">
              <div className="px-6 py-4 border-b bg-gradient-to-r from-blue-50 to-indigo-50">
                <h2 className="text-2xl font-bold text-gray-900">Top {topN} Manufacturers</h2>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="px-6 py-3 text-left text-xs font-semibold text-gray-700">#</th>
                      <th className="px-6 py-3 text-left text-xs font-semibold text-gray-700">Manufacturer</th>
                      <th className="px-6 py-3 text-right text-xs font-semibold text-gray-700">2023</th>
                      <th className="px-6 py-3 text-right text-xs font-semibold text-gray-700">2024</th>
                      <th className="px-6 py-3 text-right text-xs font-semibold text-gray-700">2023→24 Change</th>
                      <th className="px-6 py-3 text-right text-xs font-semibold text-gray-700">2025</th>
                      <th className="px-6 py-3 text-right text-xs font-semibold text-gray-700">2024→25 Change</th>
                      <th className="px-6 py-3 text-right text-xs font-semibold text-gray-700">2024 GP%</th>
                    </tr>
                  </thead>
                  <tbody>
                    {topMfg.map((item) => (
                      <tr key={item.manufacturer} className="border-b hover:bg-blue-50">
                        <td className="px-6 py-4 bg-blue-50 font-bold">{item.rank}</td>
                        <td className="px-6 py-4 font-semibold">{item.manufacturer}</td>
                        <td className="px-6 py-4 text-right text-gray-600">{formatCurrency(item.revenue2023)}</td>
                        <td className="px-6 py-4 text-right font-semibold">{formatCurrency(item.revenue2024)}</td>
                        <td className="px-6 py-4 text-right">
                          <div className="flex items-center justify-end gap-1">
                            {item.change2324 >= 0 ? <TrendingUp className="h-4 w-4 text-green-600" /> : <TrendingDown className="h-4 w-4 text-red-600" />}
                            <span className={item.change2324 >= 0 ? 'text-green-600 font-semibold' : 'text-red-600 font-semibold'}>{formatCurrency(item.change2324)}</span>
                            <span className={item.change2324 >= 0 ? 'text-green-600 text-xs' : 'text-red-600 text-xs'}>({item.changePct2324.toFixed(1)}%)</span>
                          </div>
                        </td>
                        <td className="px-6 py-4 text-right text-gray-600">{formatCurrency(item.revenue2025)}</td>
                        <td className="px-6 py-4 text-right">
                          <div className="flex items-center justify-end gap-1">
                            {item.change2425 >= 0 ? <TrendingUp className="h-4 w-4 text-green-600" /> : <TrendingDown className="h-4 w-4 text-red-600" />}
                            <span className={item.change2425 >= 0 ? 'text-green-600 font-semibold text-xs' : 'text-red-600 font-semibold text-xs'}>{formatCurrency(item.change2425)}</span>
                            <span className={item.change2425 >= 0 ? 'text-green-600 text-xs' : 'text-red-600 text-xs'}>({item.changePct2425.toFixed(1)}%)</span>
                          </div>
                        </td>
                        <td className="px-6 py-4 text-right"><span className="bg-blue-100 text-blue-800 px-3 py-1 rounded text-xs font-semibold">{item.gpPct2024.toFixed(2)}%</span></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {!data && !error && (
          <div className="bg-white rounded-xl shadow-lg p-12 text-center">
            <p className="text-gray-600 text-lg">Upload an Excel file to get started</p>
          </div>
        )}
      </div>
    </div>
  );
};

export default Dashboard;