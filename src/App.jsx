import React, { useState, useMemo, useEffect } from 'react';
import { initializeApp } from 'firebase/app';
import { getAuth, signInWithCustomToken, signInAnonymously, onAuthStateChanged } from 'firebase/auth';
import { getFirestore, collection, doc, setDoc, deleteDoc, onSnapshot } from 'firebase/firestore';
import { 
  Package, ArrowDownToLine, ArrowUpFromLine, Settings, LayoutDashboard, Plus, History,
  AlertCircle, CalendarDays, BarChart as BarChartIcon, Pencil, Check, X, Trash2, Download
} from 'lucide-react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';

// --- 엑셀(XLSX) 다운로드 유틸리티 함수 ---
const loadXlsxScript = () => {
  return new Promise((resolve, reject) => {
    if (window.XLSX) {
      resolve(window.XLSX);
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.onload = () => resolve(window.XLSX);
    script.onerror = () => reject(new Error('Excel 라이브러리를 불러오는데 실패했습니다.'));
    document.head.appendChild(script);
  });
};

const exportToExcel = async (filename, rows) => {
  try {
    const XLSX = await loadXlsxScript();
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, `${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);
  } catch (error) {
    console.error("Excel Export Error:", error);
  }
};

// --- Firebase 클라우드 연동 초기화 ---
let app, auth, db, appId;
let isFirebaseConfigured = false;

try {
  if (typeof __firebase_config !== 'undefined' && __firebase_config) {
    const config = typeof __firebase_config === 'string' ? JSON.parse(__firebase_config) : __firebase_config;
    if (config && Object.keys(config).length > 0) {
      app = initializeApp(config);
      auth = getAuth(app);
      db = getFirestore(app);
      appId = typeof __app_id !== 'undefined' ? __app_id : 'default-app-id';
      isFirebaseConfigured = true;
    }
  }
} catch (error) {
  console.warn("Firebase 클라우드 초기화 실패, 로컬 모드로 전환:", error);
}

// --- 안전한 로컬 저장소 로딩 유틸리티 (기존 데이터 마이그레이션 포함) ---
const loadLocalData = (newKey, oldKeyPrefix, fallback) => {
  try {
    let val = localStorage.getItem(newKey);
    if (!val) {
      const oldCode = localStorage.getItem('inventory_sync_code') || 'DEMO123';
      val = localStorage.getItem(`${oldKeyPrefix}_${oldCode}`);
      if (val) {
        localStorage.setItem(newKey, val); // 새 키로 데이터 복사(마이그레이션)
      }
    }
    return val ? JSON.parse(val) : fallback;
  } catch {
    return fallback;
  }
};

const loadLocalString = (newKey, oldKeyPrefix, fallback) => {
  try {
    let val = localStorage.getItem(newKey);
    if(!val) {
      const oldCode = localStorage.getItem('inventory_sync_code') || 'DEMO123';
      val = localStorage.getItem(`${oldKeyPrefix}_${oldCode}`);
      if(val) {
        localStorage.setItem(newKey, val);
      }
    }
    return val || fallback;
  } catch {
    return fallback;
  }
};

// --- [탭 1] 재고 현황 대시보드 컴포넌트 ---
const DashboardTab = ({ items, transactions, inventoryStats, dashboardSort, setDashboardSort }) => (
  <div className="space-y-6">
    <div className="grid grid-cols-1 gap-4">
      <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex items-center space-x-4">
        <div className="p-3 bg-blue-100 text-blue-600 rounded-lg shrink-0">
          <Package size={24} />
        </div>
        <div>
          <p className="text-xs text-slate-500 font-medium">등록된 품목 수</p>
          <p className="text-xl font-bold text-slate-800">{items.length}개</p>
        </div>
      </div>
    </div>

    <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
      <div className="px-4 py-3 border-b border-slate-200 bg-slate-50/50 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
        <h2 className="text-base font-semibold text-slate-800">현재 재고 현황</h2>
        <div className="flex items-center gap-2 overflow-x-auto w-full sm:w-auto pb-1 sm:pb-0 [&::-webkit-scrollbar]:hidden">
          <button
            onClick={() => {
              const rows = [
                ['구분', '업체명', '품목명', '비고', '현재 재고(개)'],
                ...inventoryStats.map(item => [item.division || '', item.category || '', item.name || '', item.note || '', item.currentStock])
              ];
              exportToExcel('현재_재고_현황', rows);
            }}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-green-700 bg-green-100 hover:bg-green-200 rounded-md transition-colors whitespace-nowrap shrink-0"
          >
            <Download size={14} />
            <span className="hidden sm:inline">엑셀 다운로드</span>
          </button>
          <div className="flex bg-slate-200/50 p-1 rounded-lg shrink-0">
            <button
              onClick={() => setDashboardSort('name')}
              className={`px-3 py-1.5 text-xs font-medium rounded-md transition-all ${dashboardSort === 'name' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}
            >
              가나다순
            </button>
            <button
              onClick={() => setDashboardSort('category')}
              className={`px-3 py-1.5 text-xs font-medium rounded-md transition-all ${dashboardSort === 'category' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}
            >
              업체명순
            </button>
            <button
              onClick={() => setDashboardSort('division')}
              className={`px-3 py-1.5 text-xs font-medium rounded-md transition-all ${dashboardSort === 'division' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}
            >
              구분순
            </button>
          </div>
        </div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-left text-xs whitespace-nowrap">
          <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
            <tr>
              <th className="px-2 sm:px-4 py-2.5 font-medium">구분</th>
              <th className="px-2 sm:px-4 py-2.5 font-medium">업체명</th>
              <th className="px-2 sm:px-4 py-2.5 font-medium">품목명</th>
              <th className="px-2 sm:px-4 py-2.5 font-medium max-w-[100px] truncate">비고</th>
              <th className="px-2 sm:px-4 py-2.5 font-medium text-right">현재 재고</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200">
            {inventoryStats.length === 0 ? (
              <tr>
                <td colSpan="5" className="px-4 py-6 text-center text-slate-500">등록된 품목이 없습니다.</td>
              </tr>
            ) : (
              inventoryStats.map(item => {
                return (
                  <tr key={item.id} className="hover:bg-slate-50/50 transition-colors">
                    <td className="px-2 sm:px-4 py-2 text-slate-500">
                      <span className="px-1.5 py-0.5 bg-slate-100 rounded text-[10px]">{item.division || '-'}</span>
                    </td>
                    <td className="px-2 sm:px-4 py-2 text-slate-500 truncate max-w-[80px] sm:max-w-[120px]" title={item.category}>{item.category}</td>
                    <td className="px-2 sm:px-4 py-2 font-medium text-slate-800 truncate max-w-[120px] sm:max-w-[200px]" title={item.name}>{item.name}</td>
                    <td className="px-2 sm:px-4 py-2 text-slate-500 truncate max-w-[100px]" title={item.note}>{item.note || '-'}</td>
                    <td className="px-2 sm:px-4 py-2 text-right font-bold text-blue-600 text-sm">{item.currentStock}개</td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </div>
  </div>
);

// --- [탭 2] 입출고 등록 컴포넌트 ---
const TransactionTab = ({ items, transactions, onAddTransaction, inventoryStats }) => {
  const today = new Date().toISOString().split('T')[0];
  const [formData, setFormData] = useState({ 
    itemId: '', 
    type: 'IN', 
    quantity: '', 
    note: '',
    date: today
  });
  const [error, setError] = useState('');
  const [successMsg, setSuccessMsg] = useState('');

  const sortedItems = useMemo(() => {
    return [...items].sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ko-KR'));
  }, [items]);

  const sortedTransactions = useMemo(() => {
    return [...transactions].sort((a, b) => {
      const dateA = a.date || '';
      const dateB = b.date || '';
      if (dateA !== dateB) return dateB.localeCompare(dateA); 
      return new Date(b.createdAt || 0) - new Date(a.createdAt || 0); 
    });
  }, [transactions]);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setSuccessMsg('');

    if (!formData.itemId || !formData.quantity || !formData.date) {
      setError('일자, 품목, 수량을 모두 입력해주세요.');
      return;
    }

    const qty = parseInt(formData.quantity, 10);
    if (qty <= 0) {
      setError('수량은 1 이상이어야 합니다.');
      return;
    }

    if (formData.type === 'OUT') {
      const itemStat = inventoryStats.find(i => i.id === formData.itemId);
      if (itemStat && itemStat.currentStock < qty) {
        setError(`현재 재고(${itemStat.currentStock}개)보다 많은 수량을 출고할 수 없습니다.`);
        return;
      }
    }

    const newTx = {
      id: `t${Date.now()}`,
      itemId: formData.itemId,
      type: formData.type,
      quantity: qty,
      date: formData.date, 
      note: formData.note,
      createdAt: new Date().toISOString()
    };
    
    await onAddTransaction(newTx);
    
    setFormData({ itemId: '', type: 'IN', quantity: '', note: '', date: today }); 
    setSuccessMsg('입출고 내역이 성공적으로 저장되었습니다!');
    setTimeout(() => setSuccessMsg(''), 3000);
  };

  return (
    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <div className="lg:col-span-1 bg-white rounded-xl border border-slate-200 shadow-sm p-4 sm:p-5 h-fit">
        <h2 className="text-base font-semibold text-slate-800 mb-4 flex items-center gap-2">
          <Plus size={18} className="text-blue-600" />
          입출고 등록
        </h2>
        {error && (
          <div className="mb-4 p-2 bg-red-50 text-red-600 text-xs rounded-lg flex items-center gap-2">
            <AlertCircle size={14} className="shrink-0"/>{error}
          </div>
        )}
        {successMsg && (
          <div className="mb-4 p-2 bg-blue-50 text-blue-600 text-xs rounded-lg flex items-center gap-2">
            <Check size={14} className="shrink-0"/>{successMsg}
          </div>
        )}
        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">일자</label>
            <input type="date" required value={formData.date} onChange={(e) => setFormData(prev => ({ ...prev, date: e.target.value }))} className="w-full rounded border-slate-200 border px-2.5 py-1.5 text-xs focus:outline-none focus:ring-1 focus:ring-blue-500 bg-white" />
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">작업 유형</label>
            <div className="grid grid-cols-2 gap-2">
              <button type="button" onClick={() => setFormData(prev => ({ ...prev, type: 'IN' }))} className={`py-1.5 px-2 rounded border text-xs font-medium transition-all flex items-center justify-center gap-1 ${formData.type === 'IN' ? 'bg-blue-600 border-blue-600 text-white shadow-sm' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'}`}>
                {formData.type === 'IN' && <Check size={12} className="shrink-0"/>} 입고 (+)
              </button>
              <button type="button" onClick={() => setFormData(prev => ({ ...prev, type: 'OUT' }))} className={`py-1.5 px-2 rounded border text-xs font-medium transition-all flex items-center justify-center gap-1 ${formData.type === 'OUT' ? 'bg-red-600 border-red-600 text-white shadow-sm' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'}`}>
                {formData.type === 'OUT' && <Check size={12} className="shrink-0"/>} 출고 (-)
              </button>
            </div>
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">품목 선택</label>
            <select value={formData.itemId} onChange={(e) => setFormData(prev => ({ ...prev, itemId: e.target.value }))} className="w-full rounded border-slate-200 border px-2.5 py-1.5 text-xs focus:outline-none focus:ring-1 focus:ring-blue-500 bg-white truncate">
              <option value="">품목 선택</option>
              {sortedItems.map(item => (
                <option key={item.id} value={item.id}>({item.category || item.provider}) {item.name}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">수량</label>
            <input type="number" min="1" value={formData.quantity} onChange={(e) => setFormData(prev => ({ ...prev, quantity: e.target.value }))} placeholder="수량 입력" className="w-full rounded border-slate-200 border px-2.5 py-1.5 text-xs focus:outline-none focus:ring-1 focus:ring-blue-500" />
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">비고 (선택)</label>
            <input type="text" value={formData.note} onChange={(e) => setFormData(prev => ({ ...prev, note: e.target.value }))} placeholder="추가 내용" className="w-full rounded border-slate-200 border px-2.5 py-1.5 text-xs focus:outline-none focus:ring-1 focus:ring-blue-500" />
          </div>
          <button type="submit" className="w-full text-white text-xs font-medium py-2 rounded bg-slate-800 hover:bg-slate-700 transition-colors mt-2">
            저장하기
          </button>
        </form>
      </div>

      <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-200 bg-slate-50/50 flex justify-between items-center flex-wrap gap-2">
          <h2 className="text-base font-semibold text-slate-800 flex items-center gap-2">
            <History size={18} className="text-slate-500" /> 최근 내역
          </h2>
          <button
            onClick={() => {
              const rows = [['일자', '유형', '품목명', '수량', '비고'], ...sortedTransactions.map(tx => {
                const item = items.find(i => i.id === tx.itemId);
                return [tx.date || '', tx.type === 'IN' ? '입고' : '출고', item ? item.name : '삭제된품목', tx.quantity, tx.note || ''];
              })];
              exportToExcel('최근_입출고_내역', rows);
            }}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-green-700 bg-green-100 hover:bg-green-200 rounded-md transition-colors"
          >
            <Download size={14} /> <span className="hidden sm:inline">다운로드</span>
          </button>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs whitespace-nowrap">
            <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
              <tr>
                <th className="px-2 sm:px-4 py-2.5 font-medium">일자</th>
                <th className="px-2 sm:px-4 py-2.5 font-medium">유형</th>
                <th className="px-2 sm:px-4 py-2.5 font-medium">품목명</th>
                <th className="px-2 sm:px-4 py-2.5 font-medium text-right">수량</th>
                <th className="px-2 sm:px-4 py-2.5 font-medium max-w-[100px]">비고</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {sortedTransactions.length === 0 ? (
                <tr><td colSpan="5" className="px-4 py-6 text-center text-slate-500">내역이 없습니다.</td></tr>
              ) : (
                sortedTransactions.map(tx => {
                  const item = items.find(i => i.id === tx.itemId);
                  const shortDate = tx.date ? tx.date.substring(5) : '';
                  return (
                    <tr key={tx.id} className="hover:bg-slate-50/50">
                      <td className="px-2 sm:px-4 py-2 text-slate-500">
                        <span className="sm:hidden">{shortDate}</span>
                        <span className="hidden sm:inline">{tx.date}</span>
                      </td>
                      <td className="px-2 sm:px-4 py-2">
                        <span className={`px-1.5 py-0.5 rounded text-[10px] font-medium ${tx.type === 'IN' ? 'bg-blue-100 text-blue-700' : 'bg-red-100 text-red-700'}`}>
                          {tx.type === 'IN' ? '입고' : '출고'}
                        </span>
                      </td>
                      <td className="px-2 sm:px-4 py-2 font-medium truncate max-w-[100px] sm:max-w-[200px]" title={item ? item.name : '삭제된품목'}>{item ? item.name : '삭제된품목'}</td>
                      <td className={`px-2 sm:px-4 py-2 text-right font-semibold ${tx.type === 'IN' ? 'text-blue-600' : 'text-red-600'}`}>{tx.type === 'IN' ? '+' : '-'}{tx.quantity}</td>
                      <td className="px-2 sm:px-4 py-2 text-slate-500 truncate max-w-[80px] sm:max-w-[120px]" title={tx.note}>{tx.note || '-'}</td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// --- [탭 3] 일별 내역 (달력/그래프) 컴포넌트 ---
const CalendarTab = ({ items, transactions, onDeleteTransaction }) => {
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().split('T')[0]);
  const [chartPeriod, setChartPeriod] = useState('day'); 
  const [deletingTxId, setDeletingTxId] = useState(null);

  const dailyTransactions = useMemo(() => {
    return transactions.filter(tx => tx.date === selectedDate).sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
  }, [transactions, selectedDate]);

  const dailyStats = useMemo(() => {
    return {
      totalIn: dailyTransactions.filter(t => t.type === 'IN').reduce((sum, t) => sum + t.quantity, 0),
      totalOut: dailyTransactions.filter(t => t.type === 'OUT').reduce((sum, t) => sum + t.quantity, 0)
    };
  }, [dailyTransactions]);

  const chartData = useMemo(() => {
    const grouped = {};
    transactions.forEach(tx => {
      if (!tx.date) return;
      const [yStr, mStr, dStr] = tx.date.split('-');
      const y = parseInt(yStr, 10);
      const m = parseInt(mStr, 10);
      const d = parseInt(dStr, 10);
      const dateObj = new Date(y, m - 1, d);
      
      let key = '', display = '';
      if (chartPeriod === 'month') {
        key = `${y}-${String(m).padStart(2, '0')}`;
        display = `${y}년 ${m}월`;
      } else if (chartPeriod === 'week') {
        const day = dateObj.getDay();
        const diff = d - day + (day === 0 ? -6 : 1); 
        const monday = new Date(y, m - 1, diff);
        key = `${monday.getFullYear()}-${String(monday.getMonth() + 1).padStart(2, '0')}-${String(monday.getDate()).padStart(2, '0')}`;
        display = `${monday.getMonth() + 1}/${monday.getDate()} 주차`;
      } else {
        key = tx.date;
        display = `${m}/${d}`;
      }

      if (!grouped[key]) grouped[key] = { key, name: display, IN: 0, OUT: 0, sortDate: dateObj.getTime() };
      if (tx.type === 'IN') grouped[key].IN += tx.quantity;
      if (tx.type === 'OUT') grouped[key].OUT += tx.quantity;
    });

    return Object.values(grouped).sort((a, b) => a.sortDate - b.sortDate).slice(-15); 
  }, [transactions, chartPeriod]);

  return (
    <div className="space-y-5">
      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4 gap-3">
          <h2 className="text-base font-semibold text-slate-800 flex items-center gap-2">
            <BarChartIcon size={18} className="text-blue-600" /> 입출고 추이
          </h2>
          <div className="flex bg-slate-100 p-1 rounded-lg w-full sm:w-auto">
            {['day', 'week', 'month'].map(period => (
              <button key={period} onClick={() => setChartPeriod(period)} className={`flex-1 sm:flex-none px-3 py-1.5 text-xs font-medium rounded transition-all ${chartPeriod === period ? 'bg-white shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}>
                {period === 'day' ? '일단위' : period === 'week' ? '주단위' : '월단위'}
              </button>
            ))}
          </div>
        </div>
        
        <div className="w-full" style={{ height: '250px' }}>
          {chartData.length > 0 ? (
            <ResponsiveContainer width="100%" height={250}>
              <BarChart data={chartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} dy={10} />
                <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                <Tooltip cursor={{ fill: '#f1f5f9' }} contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '11px' }} />
                <Legend iconType="circle" wrapperStyle={{ paddingTop: '10px', fontSize: '11px' }} />
                <Bar dataKey="IN" name="입고량" fill="#3b82f6" radius={[4, 4, 0, 0]} barSize={12} />
                <Bar dataKey="OUT" name="출고량" fill="#ef4444" radius={[4, 4, 0, 0]} barSize={12} />
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <div className="h-full flex items-center justify-center text-slate-400 text-xs">표시할 데이터가 없습니다.</div>
          )}
        </div>
      </div>

      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
        <h2 className="text-base font-semibold flex items-center gap-2"><CalendarDays size={18} className="text-blue-600 shrink-0" /> 상세 날짜 선택</h2>
        <input type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} className="w-full sm:w-auto rounded border px-3 py-1.5 text-xs focus:ring-1 focus:ring-blue-500" />
      </div>

      <div className="grid grid-cols-2 gap-3 sm:gap-4">
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row items-start sm:items-center gap-3">
          <div className="p-2 bg-blue-100 text-blue-600 rounded-lg shrink-0">
            <ArrowDownToLine size={20} className="sm:w-5 sm:h-5"/>
          </div>
          <div>
            <p className="text-xs text-slate-500 font-medium">일일 입고</p>
            <p className="text-base sm:text-xl font-bold text-slate-800">{dailyStats.totalIn}개</p>
          </div>
        </div>
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col sm:flex-row items-start sm:items-center gap-3">
          <div className="p-2 bg-red-100 text-red-600 rounded-lg shrink-0">
            <ArrowUpFromLine size={20} className="sm:w-5 sm:h-5"/>
          </div>
          <div>
            <p className="text-xs text-slate-500 font-medium">일일 출고</p>
            <p className="text-base sm:text-xl font-bold text-slate-800">{dailyStats.totalOut}개</p>
          </div>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-200 bg-slate-50/50">
          <h2 className="text-sm sm:text-base font-semibold">해당 일자 입출고 상세</h2>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs whitespace-nowrap">
            <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
              <tr>
                <th className="px-2 sm:px-4 py-2.5 font-medium text-center w-12">번호</th>
                <th className="px-2 sm:px-4 py-2.5 font-medium">유형</th>
                <th className="px-2 sm:px-4 py-2.5 font-medium">품목명</th>
                <th className="px-2 sm:px-4 py-2.5 font-medium text-right">수량</th>
                <th className="px-2 sm:px-4 py-2.5 font-medium max-w-[100px]">비고</th>
                <th className="px-2 sm:px-4 py-2.5 font-medium text-center">관리</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {dailyTransactions.length === 0 ? (
                <tr><td colSpan="6" className="px-4 py-6 text-center text-slate-500">기록된 내역이 없습니다.</td></tr>
              ) : (
                dailyTransactions.map((tx, index) => {
                  const item = items.find(i => i.id === tx.itemId);
                  // 배열이 최신순(내림차순) 정렬이므로, 역순으로 번호를 부여하여 오래된 내역부터 1번으로 시작하도록 함
                  const seqNum = dailyTransactions.length - index; 
                  return (
                    <tr key={tx.id} className="hover:bg-slate-50/50">
                      {deletingTxId === tx.id ? (
                        <>
                          <td colSpan="4" className="p-2 text-center text-red-600 font-bold text-xs">이 내역을 삭제할까요?</td>
                          <td colSpan="2" className="p-2 text-center whitespace-nowrap">
                            <button onClick={() => { onDeleteTransaction(tx.id); setDeletingTxId(null); }} className="px-2 py-1 bg-red-600 text-white text-[10px] rounded mr-1">네</button>
                            <button onClick={() => setDeletingTxId(null)} className="px-2 py-1 bg-slate-200 text-[10px] rounded">아니오</button>
                          </td>
                        </>
                      ) : (
                        <>
                          <td className="px-2 sm:px-4 py-2 text-slate-500 text-center">{seqNum}</td>
                          <td className="px-2 sm:px-4 py-2">
                            <span className={`px-1.5 py-0.5 rounded text-[10px] font-medium ${tx.type === 'IN' ? 'bg-blue-100 text-blue-700' : 'bg-red-100 text-red-700'}`}>
                              {tx.type === 'IN' ? '입고' : '출고'}
                            </span>
                          </td>
                          <td className="px-2 sm:px-4 py-2 font-medium truncate max-w-[120px] sm:max-w-[200px]" title={item ? item.name : '삭제된품목'}>{item ? item.name : '삭제된품목'}</td>
                          <td className={`px-2 sm:px-4 py-2 text-right font-semibold ${tx.type === 'IN' ? 'text-blue-600' : 'text-red-600'}`}>{tx.type === 'IN' ? '+' : '-'}{tx.quantity}</td>
                          <td className="px-2 sm:px-4 py-2 text-slate-500 truncate max-w-[80px] sm:max-w-[120px]" title={tx.note}>{tx.note || '-'}</td>
                          <td className="px-2 sm:px-4 py-2 text-center whitespace-nowrap">
                            <button onClick={() => setDeletingTxId(tx.id)} className="p-1.5 text-slate-400 hover:text-red-600" title="내역 삭제">
                              <Trash2 size={14}/>
                            </button>
                          </td>
                        </>
                      )}
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// --- [탭 4] 품목 관리 컴포넌트 ---
const ItemManagementTab = ({ items, onAddItem, onUpdateItem, onDeleteItem }) => {
  const [newItem, setNewItem] = useState({ division: '', name: '', category: '', note: '' });
  const [editingItemId, setEditingItemId] = useState(null);
  const [editFormData, setEditFormData] = useState({});
  const [deletingItemId, setDeletingItemId] = useState(null);

  const sortedItems = useMemo(() => [...items].sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ko-KR')), [items]);

  const handleAdd = async (e) => {
    e.preventDefault();
    if (!newItem.name) return;
    await onAddItem({ id: `i${Date.now()}`, ...newItem, createdAt: new Date().toISOString() });
    setNewItem({ division: '', name: '', category: '', note: '' });
  };

  const handleSaveEdit = async (id) => {
    if (!editFormData.name) return;
    await onUpdateItem({ ...items.find(i => i.id === id), ...editFormData, updatedAt: new Date().toISOString() });
    setEditingItemId(null); 
  };

  return (
    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <div className="lg:col-span-1 bg-white rounded-xl border border-slate-200 shadow-sm p-4 sm:p-5 h-fit">
        <h2 className="text-base font-semibold text-slate-800 mb-4 flex items-center gap-2"><Plus size={18} className="text-blue-600" /> 새 품목 등록</h2>
        <form onSubmit={handleAdd} className="space-y-3">
          <input type="text" value={newItem.division} onChange={e => setNewItem({...newItem, division: e.target.value})} placeholder="구분" className="w-full border p-2 rounded text-xs" />
          <input type="text" value={newItem.category} onChange={e => setNewItem({...newItem, category: e.target.value})} placeholder="업체명" className="w-full border p-2 rounded text-xs" />
          <input type="text" required value={newItem.name} onChange={e => setNewItem({...newItem, name: e.target.value})} placeholder="품목명 (필수)" className="w-full border p-2 rounded text-xs" />
          <input type="text" value={newItem.note} onChange={e => setNewItem({...newItem, note: e.target.value})} placeholder="비고" className="w-full border p-2 rounded text-xs" />
          <button type="submit" className="w-full bg-slate-800 text-white p-2 rounded hover:bg-slate-700 text-xs font-medium">품목 추가하기</button>
        </form>
      </div>

      <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="px-4 py-3 border-b bg-slate-50 flex justify-between items-center flex-wrap gap-2">
          <h2 className="text-base font-semibold flex items-center gap-2"><Settings size={18} className="text-slate-500" /> 등록 품목</h2>
          <div className="flex items-center gap-2">
            <button onClick={() => {
              const rows = [['구분', '업체명', '품목명', '비고'], ...sortedItems.map(i => [i.division, i.category, i.name, i.note])];
              exportToExcel('품목목록', rows);
            }} className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs bg-green-100 text-green-700 rounded transition-colors"><Download size={12} /><span className="hidden sm:inline">엑셀 다운로드</span></button>
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs whitespace-nowrap">
            <thead className="bg-slate-50 text-slate-600 border-b">
              <tr>
                <th className="p-2 sm:p-3 font-medium">구분</th>
                <th className="p-2 sm:p-3 font-medium">업체명</th>
                <th className="p-2 sm:p-3 font-medium">품목명</th>
                <th className="p-2 sm:p-3 font-medium max-w-[80px]">비고</th>
                <th className="p-2 sm:p-3 font-medium text-center">관리</th>
              </tr>
            </thead>
            <tbody className="divide-y">
              {sortedItems.length === 0 ? (
                <tr>
                  <td colSpan="5" className="p-4 text-center text-slate-500">등록된 품목이 없습니다.</td>
                </tr>
              ) : (
                sortedItems.map(item => {
                  return (
                    <tr key={item.id} className="hover:bg-slate-50">
                      {editingItemId === item.id ? (
                        <>
                          <td className="p-1 sm:p-2"><input className="border p-1 w-full text-xs" value={editFormData.division} onChange={e=>setEditFormData({...editFormData, division: e.target.value})}/></td>
                          <td className="p-1 sm:p-2"><input className="border p-1 w-full text-xs" value={editFormData.category} onChange={e=>setEditFormData({...editFormData, category: e.target.value})}/></td>
                          <td className="p-1 sm:p-2"><input className="border p-1 w-full text-xs" value={editFormData.name} onChange={e=>setEditFormData({...editFormData, name: e.target.value})}/></td>
                          <td className="p-1 sm:p-2"><input className="border p-1 w-full text-xs" value={editFormData.note} onChange={e=>setEditFormData({...editFormData, note: e.target.value})}/></td>
                          <td className="p-1 sm:p-2 text-center whitespace-nowrap">
                            <button onClick={() => handleSaveEdit(item.id)} className="p-1 text-blue-600 bg-blue-100 rounded mr-1"><Check size={14}/></button>
                            <button onClick={() => setEditingItemId(null)} className="p-1 text-slate-600 bg-slate-200 rounded"><X size={14}/></button>
                          </td>
                        </>
                      ) : deletingItemId === item.id ? (
                        <>
                          <td colSpan="3" className="p-2 text-center text-red-600 font-bold text-xs">삭제할까요?</td>
                          <td></td>
                          <td className="p-2 text-center whitespace-nowrap">
                            <button onClick={() => { onDeleteItem(item.id); setDeletingItemId(null); }} className="px-2 py-1 bg-red-600 text-white text-[10px] rounded mr-1">네</button>
                            <button onClick={() => setDeletingItemId(null)} className="px-2 py-1 bg-slate-200 text-[10px] rounded">아니오</button>
                          </td>
                        </>
                      ) : (
                        <>
                          <td className="p-2 sm:p-3"><span className="bg-slate-100 px-1.5 py-0.5 rounded text-[10px] text-slate-600">{item.division || '-'}</span></td>
                          <td className="p-2 sm:p-3 text-slate-600 truncate max-w-[60px] sm:max-w-[100px]" title={item.category}>{item.category}</td>
                          <td className="p-2 sm:p-3 font-medium text-slate-800 truncate max-w-[100px] sm:max-w-[180px]" title={item.name}>{item.name}</td>
                          <td className="p-2 sm:p-3 text-slate-400 truncate max-w-[80px]" title={item.note}>{item.note || '-'}</td>
                          <td className="p-2 sm:p-3 text-center whitespace-nowrap">
                            <button onClick={() => { setEditingItemId(item.id); setEditFormData(item); }} className="p-1.5 text-slate-400 hover:text-blue-600"><Pencil size={14}/></button>
                            <button onClick={() => setDeletingItemId(item.id)} className="p-1.5 text-slate-400 hover:text-red-600"><Trash2 size={14}/></button>
                          </td>
                        </>
                      )}
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// --- [메인 App 컴포넌트] ---
export default function App() {
  const [activeTab, setActiveTab] = useState('dashboard');
  const [dashboardSort, setDashboardSort] = useState('name'); 
  
  // 단일 저장소 구조 (연동 코드 없이 기본값 사용)
  const [appName, setAppName] = useState(() => loadLocalString('inventory_app_name', 'inventory_app_name', '심플 재고관리'));
  const [isEditingName, setIsEditingName] = useState(false);
  const [tempAppName, setTempAppName] = useState('');

  const [items, setItems] = useState(() => loadLocalData('inventory_items', 'inventory_items', []));
  const [transactions, setTransactions] = useState(() => loadLocalData('inventory_txs', 'txs', []));
  
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);

  // Firebase 인증 처리
  useEffect(() => {
    if (!isFirebaseConfigured) {
      setAuthLoading(false);
      return;
    }
    const initAuth = async () => {
      try {
        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
          await signInWithCustomToken(auth, __initial_auth_token);
        } else {
          await signInAnonymously(auth);
        }
      } catch (error) {
        console.error("Auth failed:", error);
        setAuthLoading(false);
      }
    };
    initAuth();

    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      setAuthLoading(false);
    });
    return () => unsubscribe();
  }, []);

  // 실시간 클라우드 구독 & 로컬 데이터 백업 (고정된 단일 경로 사용)
  useEffect(() => {
    if (!isFirebaseConfigured || !user) return;

    const itemsRef = collection(db, 'artifacts', appId, 'public', 'data', 'items');
    const unsubItems = onSnapshot(itemsRef, (snapshot) => {
      const data = snapshot.docs.map(doc => doc.data());
      setItems(data);
      try { localStorage.setItem('inventory_items', JSON.stringify(data)); } catch(e){}
    }, (error) => console.error("Items Error:", error));

    const txRef = collection(db, 'artifacts', appId, 'public', 'data', 'txs');
    const unsubTx = onSnapshot(txRef, (snapshot) => {
      const data = snapshot.docs.map(doc => doc.data());
      setTransactions(data);
      try { localStorage.setItem('inventory_txs', JSON.stringify(data)); } catch(e){}
    }, (error) => console.error("Tx Error:", error));

    const settingsRef = collection(db, 'artifacts', appId, 'public', 'data', 'settings');
    const unsubSettings = onSnapshot(settingsRef, (snapshot) => {
      const appNameDoc = snapshot.docs.find(d => d.id === 'general');
      if (appNameDoc && appNameDoc.data().appName) {
        setAppName(appNameDoc.data().appName);
        try { localStorage.setItem('inventory_app_name', appNameDoc.data().appName); } catch(e){}
      }
    }, (error) => console.error("Settings Error:", error));

    return () => {
      unsubItems();
      unsubTx();
      unsubSettings();
    };
  }, [user]);

  // 데이터 추가/수정/삭제
  const handleNameSave = async () => {
    if (tempAppName.trim() !== '') {
      const newName = tempAppName.trim();
      setAppName(newName); 
      try { localStorage.setItem('inventory_app_name', newName); } catch(e){}
      
      if (isFirebaseConfigured && user) {
        try { await setDoc(doc(db, 'artifacts', appId, 'public', 'data', 'settings', 'general'), { appName: newName }, { merge: true }); }
        catch(e) { console.error(e); }
      }
    }
    setIsEditingName(false);
  };

  const handleAddItem = async (newItem) => {
    setItems(prev => {
      const next = [...prev, newItem];
      try { localStorage.setItem('inventory_items', JSON.stringify(next)); } catch(e){}
      return next;
    });
    if (isFirebaseConfigured && user) {
      try { await setDoc(doc(db, 'artifacts', appId, 'public', 'data', 'items', newItem.id), newItem); }
      catch(e) { console.error(e); }
    }
  };

  const handleUpdateItem = async (updatedItem) => {
    setItems(prev => {
      const next = prev.map(i => i.id === updatedItem.id ? updatedItem : i);
      try { localStorage.setItem('inventory_items', JSON.stringify(next)); } catch(e){}
      return next;
    });
    if (isFirebaseConfigured && user) {
      try { await setDoc(doc(db, 'artifacts', appId, 'public', 'data', 'items', updatedItem.id), updatedItem); }
      catch(e) { console.error(e); }
    }
  };

  const handleDeleteItem = async (id) => {
    setItems(prev => {
      const next = prev.filter(i => i.id !== id);
      try { localStorage.setItem('inventory_items', JSON.stringify(next)); } catch(e){}
      return next;
    });
    if (isFirebaseConfigured && user) {
      try { await deleteDoc(doc(db, 'artifacts', appId, 'public', 'data', 'items', id)); }
      catch(e) { console.error(e); }
    }
  };

  const handleAddTransaction = async (newTx) => {
    setTransactions(prev => {
      const next = [...prev, newTx];
      try { localStorage.setItem('inventory_txs', JSON.stringify(next)); } catch(e){}
      return next;
    });
    if (isFirebaseConfigured && user) {
      try { await setDoc(doc(db, 'artifacts', appId, 'public', 'data', 'txs', newTx.id), newTx); }
      catch(e) { console.error(e); }
    }
  };

  const handleDeleteTransaction = async (id) => {
    setTransactions(prev => {
      const next = prev.filter(tx => tx.id !== id);
      try { localStorage.setItem('inventory_txs', JSON.stringify(next)); } catch(e){}
      return next;
    });
    if (isFirebaseConfigured && user) {
      try { await deleteDoc(doc(db, 'artifacts', appId, 'public', 'data', 'txs', id)); }
      catch(e) { console.error(e); }
    }
  };

  // 재고 통계 계산
  const inventoryStats = useMemo(() => {
    const stats = items.map(item => {
      const itemTx = transactions.filter(t => t.itemId === item.id);
      const totalIn = itemTx.filter(t => t.type === 'IN').reduce((sum, t) => sum + t.quantity, 0);
      const totalOut = itemTx.filter(t => t.type === 'OUT').reduce((sum, t) => sum + t.quantity, 0);
      return { ...item, currentStock: totalIn - totalOut };
    });

    return stats.sort((a, b) => {
      if (dashboardSort === 'division') {
        const divDiff = (a.division || '').localeCompare(b.division || '', 'ko-KR');
        if (divDiff !== 0) return divDiff;
      }
      if (dashboardSort === 'category' || dashboardSort === 'division') {
        const catDiff = (a.category || '').localeCompare(b.category || '', 'ko-KR');
        if (catDiff !== 0) return catDiff;
      }
      return (a.name || '').localeCompare(b.name || '', 'ko-KR'); 
    });
  }, [items, transactions, dashboardSort]);

  if (authLoading) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center">
        <div className="text-slate-500 font-medium animate-pulse flex flex-col items-center gap-2">
          <Package size={32} className="text-blue-400" /> 데이터 준비 중...
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10 shadow-sm">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-3 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex flex-col sm:flex-row justify-between w-full md:w-auto gap-4">
            <div className="flex items-center gap-2">
              <div className="bg-blue-600 p-2 rounded-lg shrink-0">
                <Package size={20} className="text-white" />
              </div>
              {isEditingName ? (
                <input
                  type="text"
                  value={tempAppName}
                  onChange={(e) => setTempAppName(e.target.value)}
                  onBlur={handleNameSave}
                  onKeyDown={(e) => e.key === 'Enter' && handleNameSave()}
                  autoFocus
                  className="text-lg font-bold text-slate-800 border-b-2 border-blue-500 outline-none w-32 sm:w-48"
                />
              ) : (
                <h1 
                  className="text-lg font-bold text-slate-800 cursor-pointer hover:text-blue-600 flex items-center gap-2 group"
                  onClick={() => { setTempAppName(appName); setIsEditingName(true); }}
                >
                  {appName} <Pencil size={14} className="text-slate-400 opacity-0 group-hover:opacity-100" />
                </h1>
              )}
            </div>
          </div>

          <nav className="flex space-x-2 overflow-x-auto w-full md:w-auto pb-1 md:pb-0 [&::-webkit-scrollbar]:hidden">
            <button onClick={() => setActiveTab('dashboard')} className={`px-3 py-1.5 rounded-md text-xs font-medium flex items-center gap-1.5 shrink-0 ${activeTab === 'dashboard' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50'}`}><LayoutDashboard size={14} />현황</button>
            <button onClick={() => setActiveTab('transactions')} className={`px-3 py-1.5 rounded-md text-xs font-medium flex items-center gap-1.5 shrink-0 ${activeTab === 'transactions' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50'}`}><ArrowDownToLine size={14} />입출고</button>
            <button onClick={() => setActiveTab('calendar')} className={`px-3 py-1.5 rounded-md text-xs font-medium flex items-center gap-1.5 shrink-0 ${activeTab === 'calendar' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50'}`}><CalendarDays size={14} />내역/그래프</button>
            <button onClick={() => setActiveTab('items')} className={`px-3 py-1.5 rounded-md text-xs font-medium flex items-center gap-1.5 shrink-0 ${activeTab === 'items' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50'}`}><Settings size={14} />품목 관리</button>
          </nav>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {activeTab === 'dashboard' && <DashboardTab items={items} transactions={transactions} inventoryStats={inventoryStats} dashboardSort={dashboardSort} setDashboardSort={setDashboardSort} />}
        {activeTab === 'transactions' && <TransactionTab items={items} transactions={transactions} onAddTransaction={handleAddTransaction} inventoryStats={inventoryStats} />}
        {activeTab === 'calendar' && <CalendarTab items={items} transactions={transactions} onDeleteTransaction={handleDeleteTransaction} />}
        {activeTab === 'items' && <ItemManagementTab items={items} onAddItem={handleAddItem} onUpdateItem={handleUpdateItem} onDeleteItem={handleDeleteItem} />}
      </main>
    </div>
  );
}