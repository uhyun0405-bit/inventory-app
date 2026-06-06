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
    // 배열 데이터를 엑셀 시트 형식으로 변환
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    // 새로운 엑셀 워크북(파일) 생성
    const workbook = XLSX.utils.book_new();
    // 워크북에 시트 추가
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    // 파일 다운로드 실행 (확장자 .xlsx)
    XLSX.writeFile(workbook, `${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);
  } catch (error) {
    console.error(error);
  }
};

// --- Firebase 클라우드 연동 초기화 ---
// 오류 발생(흰 화면)을 방지하기 위해 안전하게 초기화합니다.
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
  console.warn("Firebase 클라우드 초기화에 실패하여 로컬 모드로 전환합니다:", error);
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
        <table className="w-full text-left text-sm">
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
    
    // 로컬 상태 변경 대신 상위 컴포넌트의 클라우드 저장 함수 호출
    await onAddTransaction(newTx);
    
    setFormData({ 
      itemId: '', 
      type: 'IN', 
      quantity: '', 
      note: '',
      date: today 
    }); 
    
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
            <AlertCircle size={16} />
            {error}
          </div>
        )}
        {successMsg && (
          <div className="mb-4 p-3 bg-blue-50 text-blue-600 text-sm rounded-lg flex items-center gap-2">
            <Check size={16} />
            {successMsg}
          </div>
        )}
        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">일자</label>
            <input
              type="date"
              required
              value={formData.date}
              onChange={(e) => setFormData(prev => ({ ...prev, date: e.target.value }))}
              className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">작업 유형</label>
            <div className="grid grid-cols-2 gap-2">
              <button
                type="button"
                onClick={(e) => {
                  e.preventDefault();
                  setFormData(prev => ({ ...prev, type: 'IN' }));
                }}
                className={`py-2 px-4 rounded-lg border text-sm font-medium transition-all flex items-center justify-center gap-2 cursor-pointer ${
                  formData.type === 'IN' 
                    ? 'bg-blue-600 border-blue-600 text-white shadow-md ring-2 ring-blue-600 ring-offset-2' 
                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                }`}
              >
                {formData.type === 'IN' && <Check size={16} />}
                입고 (+)
              </button>
              <button
                type="button"
                onClick={(e) => {
                  e.preventDefault();
                  setFormData(prev => ({ ...prev, type: 'OUT' }));
                }}
                className={`py-2 px-4 rounded-lg border text-sm font-medium transition-all flex items-center justify-center gap-2 cursor-pointer ${
                  formData.type === 'OUT' 
                    ? 'bg-red-600 border-red-600 text-white shadow-md ring-2 ring-red-600 ring-offset-2' 
                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                }`}
              >
                {formData.type === 'OUT' && <Check size={16} />}
                출고 (-)
              </button>
            </div>
          </div>
          
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">품목 선택</label>
            <select
              value={formData.itemId}
              onChange={(e) => setFormData(prev => ({ ...prev, itemId: e.target.value }))}
              className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 cursor-pointer bg-white"
            >
              <option value="">품목을 선택하세요 (가나다 순)</option>
              {sortedItems.map(item => (
                <option key={item.id} value={item.id}>({item.category || item.provider}) {item.name}</option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">수량</label>
            <input
              type="number"
              min="1"
              value={formData.quantity}
              onChange={(e) => setFormData(prev => ({ ...prev, quantity: e.target.value }))}
              placeholder="수량 입력"
              className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">비고 (선택)</label>
            <input
              type="text"
              value={formData.note}
              onChange={(e) => setFormData(prev => ({ ...prev, note: e.target.value }))}
              placeholder="ex) 추가 매입, 영업팀 지급 등"
              className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>

          <button
            type="submit"
            className="w-full text-white font-medium py-2.5 rounded-lg bg-slate-800 hover:bg-slate-700 transition-colors flex justify-center items-center gap-2 mt-2"
          >
            저장하기
          </button>
        </form>
      </div>

      <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/50 flex justify-between items-center flex-wrap gap-2">
          <h2 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
            <History size={20} className="text-slate-500" />
            최근 입출고 내역
          </h2>
          <button
            onClick={() => {
              const rows = [
                ['일자', '유형', '품목명', '수량', '비고'],
                ...sortedTransactions.map(tx => {
                  const item = items.find(i => i.id === tx.itemId);
                  const typeStr = tx.type === 'IN' ? '입고' : '출고';
                  const itemName = item ? item.name : '삭제된품목';
                  return [tx.date || '', typeStr, itemName, tx.quantity, tx.note || ''];
                })
              ];
              exportToExcel('최근_입출고_내역', rows);
            }}
            className="flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-green-700 bg-green-100 hover:bg-green-200 rounded-md transition-colors"
          >
            <Download size={16} />
            <span className="hidden sm:inline">엑셀 다운로드</span>
          </button>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm">
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
                <tr>
                  <td colSpan="5" className="px-6 py-8 text-center text-slate-500">내역이 없습니다.</td>
                </tr>
              ) : (
                sortedTransactions.map(tx => {
                  const item = items.find(i => i.id === tx.itemId);
                  return (
                    <tr key={tx.id} className="hover:bg-slate-50/50 transition-colors">
                      <td className="px-6 py-4 text-slate-500 font-medium">{tx.date}</td>
                      <td className="px-6 py-4">
                        <span className={`inline-flex items-center px-2 py-1 rounded-full text-xs font-medium ${
                          tx.type === 'IN' ? 'bg-blue-100 text-blue-700' : 'bg-red-100 text-red-700'
                        }`}>
                          {tx.type === 'IN' ? '입고' : '출고'}
                        </span>
                      </td>
                      <td className="px-6 py-4 font-medium text-slate-800">{item ? item.name : '삭제된품목'}</td>
                      <td className={`px-6 py-4 text-right font-semibold ${tx.type === 'IN' ? 'text-blue-600' : 'text-red-600'}`}>
                        {tx.type === 'IN' ? '+' : '-'}{tx.quantity}
                      </td>
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

// --- [탭 3] 일별 내역 (달력) 컴포넌트 ---
const CalendarTab = ({ items, transactions }) => {
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().split('T')[0]);
  const [chartPeriod, setChartPeriod] = useState('day'); 

  const dailyTransactions = useMemo(() => {
    return transactions
      .filter(tx => tx.date === selectedDate)
      .sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
  }, [transactions, selectedDate]);

  const dailyStats = useMemo(() => {
    const totalIn = dailyTransactions.filter(t => t.type === 'IN').reduce((sum, t) => sum + t.quantity, 0);
    const totalOut = dailyTransactions.filter(t => t.type === 'OUT').reduce((sum, t) => sum + t.quantity, 0);
    return { totalIn, totalOut };
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
      
      let key = '';
      let display = '';
      
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

      if (!grouped[key]) {
        grouped[key] = { 
          key, 
          display, 
          name: display, 
          IN: 0, 
          OUT: 0, 
          sortDate: dateObj.getTime() 
        };
      }
      
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
            <BarChartIcon size={20} className="text-blue-600" />
            기간별 입출고 추이
          </h2>
          <div className="flex bg-slate-100 p-1 rounded-lg w-full sm:w-auto">
            <button
              onClick={() => setChartPeriod('day')}
              className={`flex-1 sm:flex-none px-4 py-1.5 text-sm font-medium rounded-md transition-all ${chartPeriod === 'day' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}
            >
              일단위
            </button>
            <button
              onClick={() => setChartPeriod('week')}
              className={`flex-1 sm:flex-none px-4 py-1.5 text-sm font-medium rounded-md transition-all ${chartPeriod === 'week' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}
            >
              주단위
            </button>
            <button
              onClick={() => setChartPeriod('month')}
              className={`flex-1 sm:flex-none px-4 py-1.5 text-sm font-medium rounded-md transition-all ${chartPeriod === 'month' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}
            >
              월단위
            </button>
          </div>
        </div>
        
        <div className="w-full h-72 min-h-[300px]">
          {chartData.length > 0 ? (
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} />
                <Tooltip 
                  cursor={{ fill: '#f1f5f9' }}
                  contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                />
                <Legend iconType="circle" wrapperStyle={{ paddingTop: '20px' }} />
                <Bar dataKey="IN" name="입고량" fill="#3b82f6" radius={[4, 4, 0, 0]} barSize={20} />
                <Bar dataKey="OUT" name="출고량" fill="#ef4444" radius={[4, 4, 0, 0]} barSize={20} />
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <div className="h-full flex items-center justify-center text-slate-400 text-sm">
              표시할 데이터가 없습니다.
            </div>
          )}
        </div>
      </div>

      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex items-center justify-between">
        <h2 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
          <CalendarDays size={20} className="text-blue-600" />
          상세 날짜 선택
        </h2>
        <input
          type="date"
          value={selectedDate}
          onChange={(e) => setSelectedDate(e.target.value)}
          className="rounded-lg border-slate-200 border px-4 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex items-center space-x-4">
          <div className="p-3 bg-blue-100 text-blue-600 rounded-lg">
            <ArrowDownToLine size={24} />
          </div>
          <div>
            <p className="text-sm text-slate-500 font-medium">일일 총 입고</p>
            <p className="text-2xl font-bold text-slate-800">{dailyStats.totalIn}개</p>
          </div>
        </div>
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex items-center space-x-4">
          <div className="p-3 bg-red-100 text-red-600 rounded-lg">
            <ArrowUpFromLine size={24} />
          </div>
          <div>
            <p className="text-sm text-slate-500 font-medium">일일 총 출고</p>
            <p className="text-2xl font-bold text-slate-800">{dailyStats.totalOut}개</p>
          </div>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/50">
          <h2 className="text-lg font-semibold text-slate-800">해당 일자 입출고 상세</h2>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm">
            <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
              <tr>
                <th className="px-6 py-3 font-medium">시간</th>
                <th className="px-6 py-3 font-medium">유형</th>
                <th className="px-6 py-3 font-medium">품목명</th>
                <th className="px-6 py-3 font-medium text-right">수량</th>
                <th className="px-6 py-3 font-medium">비고</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {dailyTransactions.length === 0 ? (
                <tr>
                  <td colSpan="5" className="px-6 py-8 text-center text-slate-500">해당 날짜에 기록된 내역이 없습니다.</td>
                </tr>
              ) : (
                dailyTransactions.map(tx => {
                  const item = items.find(i => i.id === tx.itemId);
                  const timeString = tx.createdAt ? new Date(tx.createdAt).toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' }) : '-';
                  return (
                    <tr key={tx.id} className="hover:bg-slate-50/50 transition-colors">
                      <td className="px-6 py-4 text-slate-500">{timeString}</td>
                      <td className="px-6 py-4">
                        <span className={`inline-flex items-center px-2 py-1 rounded-full text-xs font-medium ${
                          tx.type === 'IN' ? 'bg-blue-100 text-blue-700' : 'bg-red-100 text-red-700'
                        }`}>
                          {tx.type === 'IN' ? '입고' : '출고'}
                        </span>
                      </td>
                      <td className="px-6 py-4 font-medium text-slate-800">{item ? item.name : '알수없음'}</td>
                      <td className={`px-6 py-4 text-right font-semibold ${tx.type === 'IN' ? 'text-blue-600' : 'text-red-600'}`}>
                        {tx.type === 'IN' ? '+' : '-'}{tx.quantity}
                      </td>
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
  const [error, setError] = useState('');
  const [successMsg, setSuccessMsg] = useState('');
  
  // 수정/삭제 상태 관리
  const [editingItemId, setEditingItemId] = useState(null);
  const [editFormData, setEditFormData] = useState({ division: '', name: '', category: '', note: '' });
  const [deletingItemId, setDeletingItemId] = useState(null);

  const sortedItems = useMemo(() => {
    return [...items].sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ko-KR'));
  }, [items]);

  const handleAddItem = async (e) => {
    e.preventDefault();
    setError('');
    setSuccessMsg('');

    if (!newItem.name || !newItem.category) {
      setError('품목명과 업체명을 모두 입력해주세요.');
      return;
    }
    
    const item = {
      id: `i${Date.now()}`,
      division: newItem.division || '', 
      name: newItem.name,
      category: newItem.category,
      note: newItem.note || '', 
      createdAt: new Date().toISOString()
    };
    
    await onAddItem(item);
    setNewItem({ division: '', name: '', category: '', note: '' });
    setSuccessMsg('품목이 성공적으로 추가되었습니다!');
    setTimeout(() => setSuccessMsg(''), 3000);
  };

  const handleEditClick = (item) => {
    setEditingItemId(item.id);
    setDeletingItemId(null); // 수정 시 삭제 모드 초기화
    setEditFormData({
      division: item.division || '',
      name: item.name || '',
      category: item.category || '',
      note: item.note || ''
    });
  };

  const handleCancelEdit = () => {
    setEditingItemId(null);
  };

  const handleSaveEdit = async (id) => {
    if (!editFormData.name || !editFormData.category) {
      setError('수정 시 품목명과 업체명은 필수입니다.');
      return;
    }
    
    const currentItem = items.find(i => i.id === id);
    await onUpdateItem({ ...currentItem, ...editFormData, updatedAt: new Date().toISOString() });
    
    setEditingItemId(null); 
    setSuccessMsg('품목 정보가 수정되었습니다.');
    setTimeout(() => setSuccessMsg(''), 3000);
  };

  const confirmDelete = async (id) => {
    await onDeleteItem(id);
    setDeletingItemId(null);
    setSuccessMsg('품목이 삭제되었습니다.');
    setTimeout(() => setSuccessMsg(''), 3000);
  };

  return (
    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <div className="lg:col-span-1 bg-white rounded-xl border border-slate-200 shadow-sm p-6 h-fit">
        <h2 className="text-lg font-semibold text-slate-800 mb-4 flex items-center gap-2">
          <Plus size={20} className="text-blue-600" />
          새 품목 등록
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

        <form onSubmit={handleAddItem} className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">구분</label>
            <input
              type="text"
              value={newItem.division}
              onChange={(e) => setNewItem(prev => ({ ...prev, division: e.target.value }))}
              placeholder="ex) 자산, 소모품 등"
              className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">업체명</label>
            <input
              type="text"
              value={newItem.category}
              onChange={(e) => setNewItem(prev => ({ ...prev, category: e.target.value }))}
              placeholder="ex) (주)한국제지, 대한전자 등"
              className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">품목명</label>
            <input
              type="text"
              value={newItem.name}
              onChange={(e) => setNewItem(prev => ({ ...prev, name: e.target.value }))}
              placeholder="ex) A4 용지 (1박스)"
              className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">비고 (선택)</label>
            <input
              type="text"
              value={newItem.note}
              onChange={(e) => setNewItem(prev => ({ ...prev, note: e.target.value }))}
              placeholder="ex) 매입처, 특징 등"
              className="w-full rounded-lg border-slate-200 border px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
            />
          </div>
          <button
            type="submit"
            className="w-full text-white font-medium py-2.5 rounded-lg bg-slate-800 hover:bg-slate-700 transition-colors flex justify-center items-center gap-2 mt-2"
          >
            품목 추가하기
          </button>
        </form>
      </div>

      <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200 bg-slate-50/50 flex justify-between items-center flex-wrap gap-2">
          <h2 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
            <Settings size={20} className="text-slate-500" />
            등록된 품목 목록
          </h2>
          <div className="flex items-center gap-3">
            <button
              onClick={() => {
                const rows = [
                  ['구분', '업체명', '품목명', '비고'],
                  ...sortedItems.map(item => [item.division || '', item.category || '', item.name || '', item.note || ''])
                ];
                exportToExcel('품목목록', rows);
              }}
              className="flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-green-700 bg-green-100 hover:bg-green-200 rounded-md transition-colors"
            >
              <Download size={16} />
              <span className="hidden sm:inline">엑셀 다운로드</span>
            </button>
            <span className="text-sm text-slate-500">총 {sortedItems.length}개</span>
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm">
            <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
              <tr>
                <th className="px-6 py-3 font-medium">구분</th>
                <th className="px-6 py-3 font-medium">업체명</th>
                <th className="px-6 py-3 font-medium">품목명</th>
                <th className="px-6 py-3 font-medium">비고</th>
                <th className="px-6 py-3 font-medium text-center min-w-[120px]">관리</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {sortedItems.length === 0 ? (
                <tr>
                  <td colSpan="5" className="px-6 py-8 text-center text-slate-500">등록된 품목이 없습니다.</td>
                </tr>
              ) : (
                sortedItems.map(item => (
                  <tr key={item.id} className="hover:bg-slate-50/50 transition-colors">
                    {editingItemId === item.id ? (
                      // --- [수정 모드] ---
                      <>
                        <td className="px-4 py-3">
                          <input
                            type="text"
                            value={editFormData.division}
                            onChange={(e) => setEditFormData({ ...editFormData, division: e.target.value })}
                            className="w-full rounded-md border-slate-300 border px-2 py-1.5 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                            placeholder="구분"
                          />
                        </td>
                        <td className="px-4 py-3">
                          <input
                            type="text"
                            value={editFormData.category}
                            onChange={(e) => setEditFormData({ ...editFormData, category: e.target.value })}
                            className="w-full rounded-md border-slate-300 border px-2 py-1.5 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                            placeholder="업체명"
                          />
                        </td>
                        <td className="px-4 py-3">
                          <input
                            type="text"
                            value={editFormData.name}
                            onChange={(e) => setEditFormData({ ...editFormData, name: e.target.value })}
                            className="w-full rounded-md border-slate-300 border px-2 py-1.5 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                            placeholder="품목명"
                          />
                        </td>
                        <td className="px-4 py-3">
                          <input
                            type="text"
                            value={editFormData.note}
                            onChange={(e) => setEditFormData({ ...editFormData, note: e.target.value })}
                            className="w-full rounded-md border-slate-300 border px-2 py-1.5 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                            placeholder="비고"
                          />
                        </td>
                        <td className="px-4 py-3 text-center">
                          <div className="flex items-center justify-center gap-2">
                            <button onClick={() => handleSaveEdit(item.id)} className="p-1.5 bg-blue-100 text-blue-600 rounded-md hover:bg-blue-200 transition-colors" title="저장">
                              <Check size={16} />
                            </button>
                            <button onClick={handleCancelEdit} className="p-1.5 bg-slate-100 text-slate-600 rounded-md hover:bg-slate-200 transition-colors" title="취소">
                              <X size={16} />
                            </button>
                          </div>
                        </td>
                      </>
                    ) : deletingItemId === item.id ? (
                      // --- [삭제 확인 모드] ---
                      <>
                        <td colSpan="4" className="px-6 py-4 text-center">
                          <span className="text-red-600 font-medium text-sm">정말로 이 품목을 삭제하시겠습니까?</span>
                        </td>
                        <td className="px-4 py-3 text-center">
                          <div className="flex items-center justify-center gap-2">
                            <button onClick={() => confirmDelete(item.id)} className="px-3 py-1.5 bg-red-600 text-white rounded-md hover:bg-red-700 transition-colors text-xs font-bold shadow-sm">
                              삭제
                            </button>
                            <button onClick={() => setDeletingItemId(null)} className="px-3 py-1.5 bg-slate-200 text-slate-700 rounded-md hover:bg-slate-300 transition-colors text-xs font-bold shadow-sm">
                              취소
                            </button>
                          </div>
                        </td>
                      </>
                    ) : (
                      // --- [일반 보기 모드] ---
                      <>
                        <td className="px-6 py-4 text-slate-500">
                          <span className="px-2 py-1 bg-slate-100 rounded-md text-xs">{item.division || '-'}</span>
                        </td>
                        <td className="px-6 py-4 text-slate-500">{item.category}</td>
                        <td className="px-6 py-4 font-medium text-slate-800">{item.name}</td>
                        <td className="px-6 py-4 text-slate-500 text-xs">{item.note || '-'}</td>
                        <td className="px-6 py-4 text-center">
                          <div className="flex items-center justify-center gap-1">
                            <button 
                              onClick={() => handleEditClick(item)} 
                              className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-md transition-colors"
                              title="수정"
                            >
                              <Pencil size={16} />
                            </button>
                            <button 
                              onClick={() => setDeletingItemId(item.id)} 
                              className="p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-md transition-colors"
                              title="삭제"
                            >
                              <Trash2 size={16} />
                            </button>
                          </div>
                        </td>
                      </>
                    )}
                  </tr>
                ))
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
  
  const [appName, setAppName] = useState('심플 재고관리');
  const [isEditingName, setIsEditingName] = useState(false);
  const [tempAppName, setTempAppName] = useState('');

  const [items, setItems] = useState([]);
  const [transactions, setTransactions] = useState([]);
  
  // 클라우드 연동 인증 상태
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);

  // 로컬 스토리지 데이터 로드 및 Firebase 인증 처리
  useEffect(() => {
    // 1. Firebase 설정이 없거나 실패한 경우 즉시 로컬 스토리지 모드로 작동
    if (!isFirebaseConfigured) {
      try {
        const savedItems = localStorage.getItem('inventory_items');
        if (savedItems) setItems(JSON.parse(savedItems));
        
        const savedTx = localStorage.getItem('inventory_transactions');
        if (savedTx) setTransactions(JSON.parse(savedTx));
        
        const savedName = localStorage.getItem('inventory_app_name');
        if (savedName) setAppName(savedName);
      } catch (e) {
        console.error("로컬 데이터 로드 실패:", e);
      }
      setAuthLoading(false);
      return;
    }

    // 2. Firebase가 정상일 경우 인증 및 클라우드 연동 진행
    const initAuth = async () => {
      try {
        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
          await signInWithCustomToken(auth, __initial_auth_token);
        } else {
          await signInAnonymously(auth);
        }
      } catch (error) {
        console.error("Authentication failed:", error);
        setAuthLoading(false); // 실패 시에도 무한 로딩 방지
      }
    };
    initAuth();

    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      setAuthLoading(false);
    });

    return () => unsubscribe();
  }, []);

  // --- 자동 저장 로직 (로컬 모드일 때만 실행) ---
  useEffect(() => {
    if (isFirebaseConfigured || authLoading) return;
    localStorage.setItem('inventory_items', JSON.stringify(items));
  }, [items, isFirebaseConfigured, authLoading]);

  useEffect(() => {
    if (isFirebaseConfigured || authLoading) return;
    localStorage.setItem('inventory_transactions', JSON.stringify(transactions));
  }, [transactions, isFirebaseConfigured, authLoading]);

  useEffect(() => {
    if (isFirebaseConfigured || authLoading) return;
    localStorage.setItem('inventory_app_name', appName);
  }, [appName, isFirebaseConfigured, authLoading]);


  // --- 클라우드 구독 로직 (Firebase 활성화 시에만 실행) ---
  useEffect(() => {
    if (!isFirebaseConfigured || !user) return;

    const itemsRef = collection(db, 'artifacts', appId, 'users', user.uid, 'items');
    const unsubItems = onSnapshot(itemsRef, (snapshot) => {
      const loadedItems = snapshot.docs.map(doc => doc.data());
      setItems(loadedItems);
    }, (error) => console.error("Error fetching items:", error));

    const txRef = collection(db, 'artifacts', appId, 'users', user.uid, 'transactions');
    const unsubTx = onSnapshot(txRef, (snapshot) => {
      const loadedTx = snapshot.docs.map(doc => doc.data());
      setTransactions(loadedTx);
    }, (error) => console.error("Error fetching transactions:", error));

    const settingsRef = collection(db, 'artifacts', appId, 'users', user.uid, 'settings');
    const unsubSettings = onSnapshot(settingsRef, (snapshot) => {
      const appNameDoc = snapshot.docs.find(d => d.id === 'general');
      if (appNameDoc && appNameDoc.data().appName) {
        setAppName(appNameDoc.data().appName);
      }
    }, (error) => console.error("Error fetching settings:", error));

    return () => {
      unsubItems();
      unsubTx();
      unsubSettings();
    };
  }, [user]);

  // --- 통합 데이터 쓰기/수정/삭제 (클라우드 vs 로컬 자동 판단) ---
  const handleNameSave = async () => {
    if (tempAppName.trim() !== '') {
      const newName = tempAppName.trim();
      setAppName(newName); 
      
      if (isFirebaseConfigured && user) {
        await setDoc(doc(db, 'artifacts', appId, 'users', user.uid, 'settings', 'general'), { appName: newName }, { merge: true });
      }
    }
    setIsEditingName(false);
  };

  const handleAddItem = async (newItem) => {
    if (isFirebaseConfigured && user) {
      await setDoc(doc(db, 'artifacts', appId, 'users', user.uid, 'items', newItem.id), newItem);
    } else {
      setItems(prev => [...prev, newItem]);
    }
  };

  const handleUpdateItem = async (updatedItem) => {
    if (isFirebaseConfigured && user) {
      await setDoc(doc(db, 'artifacts', appId, 'users', user.uid, 'items', updatedItem.id), updatedItem);
    } else {
      setItems(prev => prev.map(i => i.id === updatedItem.id ? updatedItem : i));
    }
  };

  const handleDeleteItem = async (id) => {
    if (isFirebaseConfigured && user) {
      await deleteDoc(doc(db, 'artifacts', appId, 'users', user.uid, 'items', id));
    } else {
      setItems(prev => prev.filter(i => i.id !== id));
    }
  };

  const handleAddTransaction = async (newTx) => {
    if (isFirebaseConfigured && user) {
      await setDoc(doc(db, 'artifacts', appId, 'users', user.uid, 'transactions', newTx.id), newTx);
    } else {
      setTransactions(prev => [...prev, newTx]);
    }
  };

  // 재고 계산 로직
  const inventoryStats = useMemo(() => {
    const stats = items.map(item => {
      const itemTransactions = transactions.filter(t => t.itemId === item.id);
      const totalIn = itemTransactions.filter(t => t.type === 'IN').reduce((sum, t) => sum + t.quantity, 0);
      const totalOut = itemTransactions.filter(t => t.type === 'OUT').reduce((sum, t) => sum + t.quantity, 0);
      const currentStock = totalIn - totalOut;

      return {
        ...item,
        currentStock
      };
    });

    if (dashboardSort === 'division') {
      return stats.sort((a, b) => {
        const divA = a.division || '';
        const divB = b.division || '';
        const divCompare = divA.localeCompare(divB, 'ko-KR');
        if (divCompare !== 0) return divCompare;
        const catA = a.category || '';
        const catB = b.category || '';
        const catCompare = catA.localeCompare(catB, 'ko-KR');
        if (catCompare !== 0) return catCompare;
        return (a.name || '').localeCompare(b.name || '', 'ko-KR'); 
      });
    } else if (dashboardSort === 'category') {
      return stats.sort((a, b) => {
        const catA = a.category || '';
        const catB = b.category || '';
        const catCompare = catA.localeCompare(catB, 'ko-KR');
        if (catCompare !== 0) return catCompare;
        return (a.name || '').localeCompare(b.name || '', 'ko-KR'); 
      });
    } else {
      return stats.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ko-KR')); 
    }
  }, [items, transactions, dashboardSort]);

  // 데이터를 불러오는 중일 때의 로딩 화면
  if (authLoading) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center">
        <div className="text-slate-500 font-medium animate-pulse flex flex-col items-center gap-2">
          <Package size={32} className="text-blue-400" />
          앱을 로드하고 있습니다...
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10 shadow-sm">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-3 sm:h-16 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 sm:gap-0">
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
                className="text-xl font-bold text-slate-800 bg-white border-b-2 border-blue-500 px-1 py-0.5 outline-none w-48"
              />
            ) : (
              <h1 
                className="text-xl font-bold text-slate-800 cursor-pointer hover:text-blue-600 transition-colors flex items-center gap-2 group"
                onClick={() => {
                  setTempAppName(appName);
                  setIsEditingName(true);
                }}
                title="클릭하여 이름 수정"
              >
                {appName}
                <Pencil size={14} className="text-slate-400 opacity-0 group-hover:opacity-100 transition-opacity" />
              </h1>
            )}
          </div>
          <nav className="flex space-x-1 sm:space-x-4 overflow-x-auto w-full sm:w-auto pb-1 sm:pb-0 [&::-webkit-scrollbar]:hidden [-ms-overflow-style:none] [scrollbar-width:none]">
            <button
              onClick={() => setActiveTab('dashboard')}
              className={`px-3 py-2 rounded-md text-sm font-medium transition-colors flex items-center gap-2 shrink-0 ${
                activeTab === 'dashboard' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50 hover:text-slate-700'
              }`}
            >
              <LayoutDashboard size={16} />
              <span className="whitespace-nowrap">재고 현황</span>
            </button>
            <button
              onClick={() => setActiveTab('transactions')}
              className={`px-3 py-2 rounded-md text-sm font-medium transition-colors flex items-center gap-2 shrink-0 ${
                activeTab === 'transactions' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50 hover:text-slate-700'
              }`}
            >
              <ArrowDownToLine size={16} />
              <span className="whitespace-nowrap">입출고 등록</span>
            </button>
            <button
              onClick={() => setActiveTab('calendar')}
              className={`px-3 py-2 rounded-md text-sm font-medium transition-colors flex items-center gap-2 shrink-0 ${
                activeTab === 'calendar' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50 hover:text-slate-700'
              }`}
            >
              <CalendarDays size={16} />
              <span className="whitespace-nowrap">일별 내역</span>
            </button>
            <button
              onClick={() => setActiveTab('items')}
              className={`px-3 py-2 rounded-md text-sm font-medium transition-colors flex items-center gap-2 shrink-0 ${
                activeTab === 'items' ? 'bg-slate-100 text-slate-900' : 'text-slate-500 hover:bg-slate-50 hover:text-slate-700'
              }`}
            >
              <Settings size={16} />
              <span className="whitespace-nowrap">품목 관리</span>
            </button>
          </nav>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {activeTab === 'dashboard' && (
          <DashboardTab 
            items={items} 
            transactions={transactions} 
            inventoryStats={inventoryStats} 
            dashboardSort={dashboardSort} 
            setDashboardSort={setDashboardSort} 
          />
        )}
        {activeTab === 'transactions' && (
          <TransactionTab 
            items={items} 
            transactions={transactions} 
            onAddTransaction={handleAddTransaction} 
            inventoryStats={inventoryStats} 
          />
        )}
        {activeTab === 'calendar' && (
          <CalendarTab 
            items={items} 
            transactions={transactions} 
          />
        )}
        {activeTab === 'items' && (
          <ItemManagementTab 
            items={items} 
            onAddItem={handleAddItem}
            onUpdateItem={handleUpdateItem}
            onDeleteItem={handleDeleteItem}
          />
        )}
      </main>
    </div>
  );
}