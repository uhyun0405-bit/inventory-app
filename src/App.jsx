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

// --- [탭 1] 재고 현황 대시보드 컴포넌트 ---
const DashboardTab = ({ items, transactions, inventoryStats, dashboardSort, setDashboardSort }) => (
  <div className="space-y-6">
    <div className="grid grid-cols-1 gap-4">
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex items-center space-x-4">
        <div className="p-3 bg-blue-100 text-blue-600 rounded-lg">
          <Package size={24} />
        </div>
        <div>
          <p className="text-sm text-slate-500 font-medium">등록된 품목 수</p>
          <p className="text-2xl font-bold text-slate-800">{items.length}개</p>
        </div>
      </div>
    </div>

    <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
      <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/50 flex justify-between items-center flex-wrap gap-3">
        <h2 className="text-lg font-semibold text-slate-800">현재 재고 현황</h2>
        <div className="flex items-center gap-2 overflow-x-auto w-full sm:w-auto">
          <button
            onClick={() => {
              const rows = [
                ['구분', '업체명', '품목명', '비고', '현재 재고(개)'],
                ...inventoryStats.map(item => [item.division || '', item.category || '', item.name || '', item.note || '', item.currentStock])
              ];
              exportToExcel('현재_재고_현황', rows);
            }}
            className="flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-green-700 bg-green-100 hover:bg-green-200 rounded-md transition-colors whitespace-nowrap"
          >
            <Download size={16} />
            <span className="hidden sm:inline">엑셀 다운로드</span>
          </button>
          <div className="flex bg-slate-200/50 p-1 rounded-lg shrink-0">
            <button
              onClick={() => setDashboardSort('name')}
              className={`px-3 py-1.5 text-sm font-medium rounded-md transition-all ${dashboardSort === 'name' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}
            >
              가나다순
            </button>
            <button
              onClick={() => setDashboardSort('category')}
              className={`px-3 py-1.5 text-sm font-medium rounded-md transition-all ${dashboardSort === 'category' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}
            >
              업체명순
            </button>
            <button
              onClick={() => setDashboardSort('division')}
              className={`px-3 py-1.5 text-sm font-medium rounded-md transition-all ${dashboardSort === 'division' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}
            >
              구분순
            </button>
          </div>
        </div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-left text-sm min-w-[600px]">
          <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
            <tr>
              <th className="px-6 py-3 font-medium">구분</th>
              <th className="px-6 py-3 font-medium">업체명</th>
              <th className="px-6 py-3 font-medium">품목명</th>
              <th className="px-6 py-3 font-medium">비고</th>
              <th className="px-6 py-3 font-medium text-right">현재 재고</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200">
            {inventoryStats.length === 0 ? (
              <tr>
                <td colSpan="5" className="px-6 py-8 text-center text-slate-500">등록된 품목이 없습니다.</td>
              </tr>
            ) : (
              inventoryStats.map(item => (
                <tr key={item.id} className="hover:bg-slate-50/50 transition-colors">
                  <td className="px-6 py-4 text-slate-500">
                    <span className="px-2 py-1 bg-slate-100 rounded-md text-xs">{item.division || '-'}</span>
                  </td>
                  <td className="px-6 py-4 text-slate-500">{item.category}</td>
                  <td className="px-6 py-4 font-medium text-slate-800">{item.name}</td>
                  <td className="px-6 py-4 text-slate-500 text-xs">{item.note || '-'}</td>
                  <td className="px-6 py-4 text-right font-bold text-blue-600">{item.currentStock}개</td>
                </tr>
              ))
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
      <div className="lg:col-span-1 bg-white rounded-xl border border-slate-200 shadow-sm p-6 h-fit">
        <h2 className="text-lg font-semibold text-slate-800 mb-4 flex items-center gap-2">
          <Plus size={20} className="text-blue-600" />
          입출고 등록
        </h2>
        {error && (
          <div className="mb-4 p-3 bg-red-50 text-red-600 text-sm rounded-lg flex items-center gap-2">
            <AlertCircle size={16} />{error}
          </div>
        )}
        {successMsg && (
          <div className="mb-4 p-3 bg-blue-50 text-blue-600 text-sm rounded-lg flex items-center gap-2">
            <Check size={16} />{successMsg}
          </div>
        )}
        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">일자</label>
            <input type="date" required value={formData.date} onChange={(e) => setFormData(prev => ({ ...prev, date: e.target.value }))} className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white" />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">작업 유형</label>
            <div className="grid grid-cols-2 gap-2">
              <button type="button" onClick={() => setFormData(prev => ({ ...prev, type: 'IN' }))} className={`py-2 px-4 rounded-lg border text-sm font-medium transition-all flex items-center justify-center gap-2 ${formData.type === 'IN' ? 'bg-blue-600 border-blue-600 text-white shadow-md ring-2 ring-blue-600 ring-offset-2' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'}`}>
                {formData.type === 'IN' && <Check size={16} />} 입고 (+)
              </button>
              <button type="button" onClick={() => setFormData(prev => ({ ...prev, type: 'OUT' }))} className={`py-2 px-4 rounded-lg border text-sm font-medium transition-all flex items-center justify-center gap-2 ${formData.type === 'OUT' ? 'bg-red-600 border-red-600 text-white shadow-md ring-2 ring-red-600 ring-offset-2' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'}`}>
                {formData.type === 'OUT' && <Check size={16} />} 출고 (-)
              </button>
            </div>
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">품목 선택</label>
            <select value={formData.itemId} onChange={(e) => setFormData(prev => ({ ...prev, itemId: e.target.value }))} className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
              <option value="">품목을 선택하세요</option>
              {sortedItems.map(item => (
                <option key={item.id} value={item.id}>({item.category || item.provider}) {item.name}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">수량</label>
            <input type="number" min="1" value={formData.quantity} onChange={(e) => setFormData(prev => ({ ...prev, quantity: e.target.value }))} placeholder="수량 입력" className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">비고 (선택)</label>
            <input type="text" value={formData.note} onChange={(e) => setFormData(prev => ({ ...prev, note: e.target.value }))} placeholder="ex) " className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <button type="submit" className="w-full text-white font-medium py-2.5 rounded-lg bg-slate-800 hover:bg-slate-700 transition-colors flex justify-center items-center gap-2 mt-2">
            저장하기
          </button>
        </form>
      </div>

      <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/50 flex justify-between items-center flex-wrap gap-2">
          <h2 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
            <History size={20} className="text-slate-500" /> 최근 입출고 내역
          </h2>
          <button
            onClick={() => {
              const rows = [['일자', '유형', '품목명', '수량', '비고'], ...sortedTransactions.map(tx => {
                const item = items.find(i => i.id === tx.itemId);
                return [tx.date || '', tx.type === 'IN' ? '입고' : '출고', item ? item.name : '삭제된품목', tx.quantity, tx.note || ''];
              })];
              exportToExcel('최근_입출고_내역', rows);
            }}
            className="flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-green-700 bg-green-100 hover:bg-green-200 rounded-md transition-colors"
          >
            <Download size={16} /> <span className="hidden sm:inline">엑셀 다운로드</span>
          </button>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm min-w-[500px]">
            <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
              <tr>
                <th className="px-6 py-3 font-medium">일자</th>
                <th className="px-6 py-3 font-medium">유형</th>
                <th className="px-6 py-3 font-medium">품목명</th>
                <th className="px-6 py-3 font-medium text-right">수량</th>
                <th className="px-6 py-3 font-medium">비고</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {sortedTransactions.length === 0 ? (
                <tr><td colSpan="5" className="px-6 py-8 text-center text-slate-500">내역이 없습니다.</td></tr>
              ) : (
                sortedTransactions.map(tx => {
                  const item = items.find(i => i.id === tx.itemId);
                  return (
                    <tr key={tx.id} className="hover:bg-slate-50/50">
                      <td className="px-6 py-4 text-slate-500">{tx.date}</td>
                      <td className="px-6 py-4">
                        <span className={`px-2 py-1 rounded-full text-xs font-medium ${tx.type === 'IN' ? 'bg-blue-100 text-blue-700' : 'bg-red-100 text-red-700'}`}>
                          {tx.type === 'IN' ? '입고' : '출고'}
                        </span>
                      </td>
                      <td className="px-6 py-4 font-medium">{item ? item.name : '삭제된품목'}</td>
                      <td className={`px-6 py-4 text-right font-semibold ${tx.type === 'IN' ? 'text-blue-600' : 'text-red-600'}`}>{tx.type === 'IN' ? '+' : '-'}{tx.quantity}</td>
                      <td className="px-6 py-4 text-slate-500">{tx.note || '-'}</td>
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
const CalendarTab = ({ items, transactions }) => {
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().split('T')[0]);
  const [chartPeriod, setChartPeriod] = useState('day'); 

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
    <div className="space-y-6">
      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-4">
          <h2 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
            <BarChartIcon size={20} className="text-blue-600" /> 기간별 입출고 추이
          </h2>
          <div className="flex bg-slate-100 p-1 rounded-lg w-full sm:w-auto">
            {['day', 'week', 'month'].map(period => (
              <button key={period} onClick={() => setChartPeriod(period)} className={`flex-1 sm:flex-none px-4 py-1.5 text-sm font-medium rounded-md ${chartPeriod === period ? 'bg-white shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}>
                {period === 'day' ? '일단위' : period === 'week' ? '주단위' : '월단위'}
              </button>
            ))}
          </div>
        </div>
        
        {/* 그래프 에러 방지: 높이를 명시적으로 300px 지정 */}
        <div className="w-full" style={{ height: '300px' }}>
          {chartData.length > 0 ? (
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={chartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                <Tooltip cursor={{ fill: '#f1f5f9' }} contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
                <Legend iconType="circle" wrapperStyle={{ paddingTop: '20px' }} />
                <Bar dataKey="IN" name="입고량" fill="#3b82f6" radius={[4, 4, 0, 0]} barSize={20} />
                <Bar dataKey="OUT" name="출고량" fill="#ef4444" radius={[4, 4, 0, 0]} barSize={20} />
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <div className="h-full flex items-center justify-center text-slate-400 text-sm">표시할 데이터가 없습니다.</div>
          )}
        </div>
      </div>

      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex items-center justify-between">
        <h2 className="text-lg font-semibold flex items-center gap-2"><CalendarDays size={20} className="text-blue-600" /> 상세 날짜 선택</h2>
        <input type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} className="rounded-lg border px-4 py-2 focus:ring-2 focus:ring-blue-500" />
      </div>

      <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/50">
          <h2 className="text-lg font-semibold">해당 일자 입출고 상세</h2>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm min-w-[500px]">
            <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
              <tr><th className="px-6 py-3 font-medium">시간</th><th className="px-6 py-3 font-medium">유형</th><th className="px-6 py-3 font-medium">품목명</th><th className="px-6 py-3 font-medium text-right">수량</th><th className="px-6 py-3 font-medium">비고</th></tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {dailyTransactions.length === 0 ? (
                <tr><td colSpan="5" className="px-6 py-8 text-center text-slate-500">기록된 내역이 없습니다.</td></tr>
              ) : (
                dailyTransactions.map(tx => {
                  const item = items.find(i => i.id === tx.itemId);
                  const timeStr = tx.createdAt ? new Date(tx.createdAt).toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' }) : '-';
                  return (
                    <tr key={tx.id} className="hover:bg-slate-50/50">
                      <td className="px-6 py-4 text-slate-500">{timeStr}</td>
                      <td className="px-6 py-4"><span className={`px-2 py-1 rounded-full text-xs font-medium ${tx.type === 'IN' ? 'bg-blue-100 text-blue-700' : 'bg-red-100 text-red-700'}`}>{tx.type === 'IN' ? '입고' : '출고'}</span></td>
                      <td className="px-6 py-4 font-medium">{item ? item.name : '알수없음'}</td>
                      <td className={`px-6 py-4 text-right font-semibold ${tx.type === 'IN' ? 'text-blue-600' : 'text-red-600'}`}>{tx.type === 'IN' ? '+' : '-'}{tx.quantity}</td>
                      <td className="px-6 py-4 text-slate-500">{tx.note || '-'}</td>
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
      <div className="lg:col-span-1 bg-white rounded-xl border border-slate-200 shadow-sm p-6 h-fit">
        <h2 className="text-lg font-semibold text-slate-800 mb-4 flex items-center gap-2"><Plus size={20} className="text-blue-600" /> 새 품목 등록</h2>
        <form onSubmit={handleAdd} className="space-y-4">
          <input type="text" value={newItem.division} onChange={e => setNewItem({...newItem, division: e.target.value})} placeholder="구분 (ex: 소모품)" className="w-full border p-2 rounded" />
          <input type="text" value={newItem.category} onChange={e => setNewItem({...newItem, category: e.target.value})} placeholder="업체명" className="w-full border p-2 rounded" />
          <input type="text" required value={newItem.name} onChange={e => setNewItem({...newItem, name: e.target.value})} placeholder="품목명 *" className="w-full border p-2 rounded" />
          <input type="text" value={newItem.note} onChange={e => setNewItem({...newItem, note: e.target.value})} placeholder="비고" className="w-full border p-2 rounded" />
          <button type="submit" className="w-full bg-slate-800 text-white p-2 rounded hover:bg-slate-700">품목 추가하기</button>
        </form>
      </div>

      <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b bg-slate-50 flex justify-between items-center flex-wrap gap-2">
          <h2 className="text-lg font-semibold flex items-center gap-2"><Settings size={20} className="text-slate-500" /> 등록된 품목 목록</h2>
          <div className="flex items-center gap-3">
            <button onClick={() => {
              const rows = [['구분', '업체명', '품목명', '비고'], ...sortedItems.map(i => [i.division, i.category, i.name, i.note])];
              exportToExcel('품목목록', rows);
            }} className="flex items-center gap-1.5 px-3 py-1.5 text-sm bg-green-100 text-green-700 rounded"><Download size={16} /><span className="hidden sm:inline">엑셀 다운로드</span></button>
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm min-w-[600px]">
            <thead className="bg-slate-50 text-slate-600 border-b">
              <tr><th className="p-3">구분</th><th className="p-3">업체명</th><th className="p-3">품목명</th><th className="p-3">비고</th><th className="p-3 text-center">관리</th></tr>
            </thead>
            <tbody className="divide-y">
              {sortedItems.map(item => (
                <tr key={item.id} className="hover:bg-slate-50">
                  {editingItemId === item.id ? (
                    <>
                      <td className="p-2"><input className="border p-1 w-full" value={editFormData.division} onChange={e=>setEditFormData({...editFormData, division: e.target.value})}/></td>
                      <td className="p-2"><input className="border p-1 w-full" value={editFormData.category} onChange={e=>setEditFormData({...editFormData, category: e.target.value})}/></td>
                      <td className="p-2"><input className="border p-1 w-full" value={editFormData.name} onChange={e=>setEditFormData({...editFormData, name: e.target.value})}/></td>
                      <td className="p-2"><input className="border p-1 w-full" value={editFormData.note} onChange={e=>setEditFormData({...editFormData, note: e.target.value})}/></td>
                      <td className="p-2 text-center">
                        <button onClick={() => handleSaveEdit(item.id)} className="p-1 text-blue-600 bg-blue-100 rounded mr-1"><Check size={16}/></button>
                        <button onClick={() => setEditingItemId(null)} className="p-1 text-slate-600 bg-slate-200 rounded"><X size={16}/></button>
                      </td>
                    </>
                  ) : deletingItemId === item.id ? (
                    <>
                      <td colSpan="4" className="p-3 text-center text-red-600 font-bold">정말 삭제하시겠습니까?</td>
                      <td className="p-2 text-center">
                        <button onClick={() => { onDeleteItem(item.id); setDeletingItemId(null); }} className="px-2 py-1 bg-red-600 text-white text-xs rounded mr-1">삭제</button>
                        <button onClick={() => setDeletingItemId(null)} className="px-2 py-1 bg-slate-200 text-xs rounded">취소</button>
                      </td>
                    </>
                  ) : (
                    <>
                      <td className="p-3"><span className="bg-slate-100 px-2 py-1 rounded text-xs">{item.division || '-'}</span></td>
                      <td className="p-3">{item.category}</td>
                      <td className="p-3 font-medium">{item.name}</td>
                      <td className="p-3 text-slate-500">{item.note}</td>
                      <td className="p-3 text-center">
                        <button onClick={() => { setEditingItemId(item.id); setEditFormData(item); }} className="p-1 text-slate-400 hover:text-blue-600"><Pencil size={16}/></button>
                        <button onClick={() => setDeletingItemId(item.id)} className="p-1 text-slate-400 hover:text-red-600"><Trash2 size={16}/></button>
                      </td>
                    </>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// --- 안전한 로컬 저장소 로딩 유틸리티 ---
const safeParse = (key, fallback) => {
  try {
    const val = localStorage.getItem(key);
    return val ? JSON.parse(val) : fallback;
  } catch {
    return fallback;
  }
};

const getInitSyncCode = () => {
  try {
    return localStorage.getItem('inventory_sync_code') || 'DEMO123';
  } catch {
    return 'DEMO123';
  }
};

// --- [메인 App 컴포넌트] ---
export default function App() {
  const [activeTab, setActiveTab] = useState('dashboard');
  const [dashboardSort, setDashboardSort] = useState('name'); 
  
  // 1. 연동 코드 캐싱 (새로고침 시 초기화 방지)
  const initCode = getInitSyncCode();
  const [syncCode, setSyncCode] = useState(initCode);
  const [syncInput, setSyncInput] = useState(initCode);

  // 2. 앱 상태 초기화 (캐시된 데이터로 즉시 화면에 띄움)
  const [appName, setAppName] = useState(() => {
    try { return localStorage.getItem(`inventory_app_name_${initCode}`) || '심플 재고관리'; }
    catch { return '심플 재고관리'; }
  });
  const [isEditingName, setIsEditingName] = useState(false);
  const [tempAppName, setTempAppName] = useState('');

  const [items, setItems] = useState(() => safeParse(`inventory_items_${initCode}`, []));
  const [transactions, setTransactions] = useState(() => safeParse(`txs_${initCode}`, []));
  
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

  // 3. 실시간 클라우드 구독 & 로컬 데이터 백업 동시 진행
  useEffect(() => {
    if (!isFirebaseConfigured || !user) return;

    const itemsRef = collection(db, 'artifacts', appId, 'public', 'data', `items_${syncCode}`);
    const unsubItems = onSnapshot(itemsRef, (snapshot) => {
      const data = snapshot.docs.map(doc => doc.data());
      setItems(data);
      try { localStorage.setItem(`inventory_items_${syncCode}`, JSON.stringify(data)); } catch(e){}
    }, (error) => console.error("Items Error:", error));

    const txRef = collection(db, 'artifacts', appId, 'public', 'data', `txs_${syncCode}`);
    const unsubTx = onSnapshot(txRef, (snapshot) => {
      const data = snapshot.docs.map(doc => doc.data());
      setTransactions(data);
      try { localStorage.setItem(`txs_${syncCode}`, JSON.stringify(data)); } catch(e){}
    }, (error) => console.error("Tx Error:", error));

    const settingsRef = collection(db, 'artifacts', appId, 'public', 'data', `settings_${syncCode}`);
    const unsubSettings = onSnapshot(settingsRef, (snapshot) => {
      const appNameDoc = snapshot.docs.find(d => d.id === 'general');
      if (appNameDoc && appNameDoc.data().appName) {
        setAppName(appNameDoc.data().appName);
        try { localStorage.setItem(`inventory_app_name_${syncCode}`, appNameDoc.data().appName); } catch(e){}
      } else {
        setAppName('심플 재고관리');
      }
    }, (error) => console.error("Settings Error:", error));

    return () => {
      unsubItems();
      unsubTx();
      unsubSettings();
    };
  }, [user, syncCode]);

  // 4. 데이터 추가/수정/삭제 시 로컬과 클라우드 모두 저장
  const handleNameSave = async () => {
    if (tempAppName.trim() !== '') {
      const newName = tempAppName.trim();
      setAppName(newName); 
      try { localStorage.setItem(`inventory_app_name_${syncCode}`, newName); } catch(e){}
      
      if (isFirebaseConfigured && user) {
        try { await setDoc(doc(db, 'artifacts', appId, 'public', 'data', `settings_${syncCode}`, 'general'), { appName: newName }, { merge: true }); }
        catch(e) { console.error(e); }
      }
    }
    setIsEditingName(false);
  };

  const handleAddItem = async (newItem) => {
    setItems(prev => {
      const next = [...prev, newItem];
      try { localStorage.setItem(`inventory_items_${syncCode}`, JSON.stringify(next)); } catch(e){}
      return next;
    });
    if (isFirebaseConfigured && user) {
      try { await setDoc(doc(db, 'artifacts', appId, 'public', 'data', `items_${syncCode}`, newItem.id), newItem); }
      catch(e) { console.error(e); }
    }
  };

  const handleUpdateItem = async (updatedItem) => {
    setItems(prev => {
      const next = prev.map(i => i.id === updatedItem.id ? updatedItem : i);
      try { localStorage.setItem(`inventory_items_${syncCode}`, JSON.stringify(next)); } catch(e){}
      return next;
    });
    if (isFirebaseConfigured && user) {
      try { await setDoc(doc(db, 'artifacts', appId, 'public', 'data', `items_${syncCode}`, updatedItem.id), updatedItem); }
      catch(e) { console.error(e); }
    }
  };

  const handleDeleteItem = async (id) => {
    setItems(prev => {
      const next = prev.filter(i => i.id !== id);
      try { localStorage.setItem(`inventory_items_${syncCode}`, JSON.stringify(next)); } catch(e){}
      return next;
    });
    if (isFirebaseConfigured && user) {
      try { await deleteDoc(doc(db, 'artifacts', appId, 'public', 'data', `items_${syncCode}`, id)); }
      catch(e) { console.error(e); }
    }
  };

  const handleAddTransaction = async (newTx) => {
    setTransactions(prev => {
      const next = [...prev, newTx];
      try { localStorage.setItem(`txs_${syncCode}`, JSON.stringify(next)); } catch(e){}
      return next;
    });
    if (isFirebaseConfigured && user) {
      try { await setDoc(doc(db, 'artifacts', appId, 'public', 'data', `txs_${syncCode}`, newTx.id), newTx); }
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
          <Package size={32} className="text-blue-400" /> 데이터 연동 준비 중...
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
                  className="text-xl font-bold text-slate-800 border-b-2 border-blue-500 outline-none w-32 sm:w-48"
                />
              ) : (
                <h1 
                  className="text-xl font-bold text-slate-800 cursor-pointer hover:text-blue-600 flex items-center gap-2 group"
                  onClick={() => { setTempAppName(appName); setIsEditingName(true); }}
                >
                  {appName} <Pencil size={14} className="text-slate-400 opacity-0 group-hover:opacity-100" />
                </h1>
              )}
            </div>

            {/* 연동 코드 입력부 */}
            <div className="flex items-center gap-2 bg-blue-50 px-3 py-1.5 rounded-lg border border-blue-200">
              <span className="text-xs font-bold text-blue-700 whitespace-nowrap">🔄 연동 코드:</span>
              <input
                type="text"
                value={syncInput}
                onChange={(e) => setSyncInput(e.target.value.toUpperCase().replace(/[^A-Z0-9_]/g, ''))}
                className="w-full sm:w-24 bg-white border border-blue-200 rounded px-1.5 py-0.5 text-sm font-bold text-slate-800 focus:ring-2 focus:ring-blue-400 outline-none"
                placeholder="코드명"
              />
              <button
                onClick={() => { 
                  const newCode = syncInput.trim();
                  if(newCode && newCode !== syncCode) {
                    setSyncCode(newCode);
                    try { localStorage.setItem('inventory_sync_code', newCode); } catch(e){}
                    setItems(safeParse(`inventory_items_${newCode}`, []));
                    setTransactions(safeParse(`txs_${newCode}`, []));
                    try { setAppName(localStorage.getItem(`inventory_app_name_${newCode}`) || '심플 재고관리'); } catch(e){}
                  }
                }}
                className="text-xs bg-blue-600 text-white px-2.5 py-1 rounded hover:bg-blue-700 font-medium"
              >
                적용
              </button>
            </div>
          </div>

          <nav className="flex space-x-2 overflow-x-auto w-full md:w-auto pb-1 md:pb-0 [&::-webkit-scrollbar]:hidden">
            <button onClick={() => setActiveTab('dashboard')} className={`px-3 py-2 rounded-md text-sm font-medium flex items-center gap-2 shrink-0 ${activeTab === 'dashboard' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50'}`}><LayoutDashboard size={16} />현황</button>
            <button onClick={() => setActiveTab('transactions')} className={`px-3 py-2 rounded-md text-sm font-medium flex items-center gap-2 shrink-0 ${activeTab === 'transactions' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50'}`}><ArrowDownToLine size={16} />입출고</button>
            <button onClick={() => setActiveTab('calendar')} className={`px-3 py-2 rounded-md text-sm font-medium flex items-center gap-2 shrink-0 ${activeTab === 'calendar' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50'}`}><CalendarDays size={16} />내역/그래프</button>
            <button onClick={() => setActiveTab('items')} className={`px-3 py-2 rounded-md text-sm font-medium flex items-center gap-2 shrink-0 ${activeTab === 'items' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50'}`}><Settings size={16} />품목 관리</button>
          </nav>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {activeTab === 'dashboard' && <DashboardTab items={items} transactions={transactions} inventoryStats={inventoryStats} dashboardSort={dashboardSort} setDashboardSort={setDashboardSort} />}
        {activeTab === 'transactions' && <TransactionTab items={items} transactions={transactions} onAddTransaction={handleAddTransaction} inventoryStats={inventoryStats} />}
        {activeTab === 'calendar' && <CalendarTab items={items} transactions={transactions} />}
        {activeTab === 'items' && <ItemManagementTab items={items} onAddItem={handleAddItem} onUpdateItem={handleUpdateItem} onDeleteItem={handleDeleteItem} />}
      </main>
    </div>
  );
}