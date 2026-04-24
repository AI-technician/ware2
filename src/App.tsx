import React, { useState, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { 
  UploadCloud, 
  MapPin, 
  Building2, 
  Maximize, 
  TableProperties,
  BarChart3,
  Search,
  FileSpreadsheet,
  Trash2,
  Filter
} from 'lucide-react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer,
  Cell
} from 'recharts';

export default function App() {
  const [data, setData] = useState<any[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  
  // Mapped Columns
  const [nameCol, setNameCol] = useState<string>('');
  const [locationCol, setLocationCol] = useState<string>('');
  const [areaCol, setAreaCol] = useState<string>('');
  const [urlCol, setUrlCol] = useState<string>('');
  const [regionCol, setRegionCol] = useState<string>('');
  
  // Filters
  const [selectedRegions, setSelectedRegions] = useState<Set<string>>(new Set());
  const [searchTerm, setSearchTerm] = useState('');
  
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const bstr = event.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const parsedData = XLSX.utils.sheet_to_json(ws);
      
      if (parsedData.length > 0) {
        const cols = Object.keys(parsedData[0] as object);
        setColumns(cols);
        setData(parsedData);
        
        // Auto-detect columns
        const findCol = (keywords: string[]) => 
          cols.find(c => keywords.some(k => c.toLowerCase().includes(k))) || cols[0];
          
        setNameCol(findCol(['상호', '창고명', '시설', '기업', '건축주', '이름', '명칭']));
        setLocationCol(findCol(['소재지', '주소', '위치']));
        setAreaCol(findCol(['면적', '연면적', '대지면적', '규모', '크기']));
        setUrlCol(findCol(['토지정보주소', 'url', '링크', '웹사이트']) || '');
        setRegionCol(findCol(['지역구분', '권역', '행정동']) || '');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (!file) return;
    
    // Check if excel
    if (!file.name.match(/\.(xlsx|xls|csv)$/)) {
      alert("엑셀 파일(.xlsx, .xls, .csv)만 업로드 가능합니다.");
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      const bstr = event.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const parsedData = XLSX.utils.sheet_to_json(ws);
      
      if (parsedData.length > 0) {
        const cols = Object.keys(parsedData[0] as object);
        setColumns(cols);
        setData(parsedData);
        
        const findCol = (keywords: string[]) => 
          cols.find(c => keywords.some(k => c.toLowerCase().includes(k))) || cols[0];
          
        setNameCol(findCol(['상호', '창고명', '시설', '기업', '건축주', '이름', '명칭']));
        setLocationCol(findCol(['소재지', '주소', '위치']));
        setAreaCol(findCol(['면적', '연면적', '대지면적', '규모', '크기']));
        setUrlCol(findCol(['토지정보주소', 'url', '링크', '웹사이트']) || '');
        setRegionCol(findCol(['지역구분', '권역', '행정동']) || '');
      }
    };
    reader.readAsBinaryString(file);
  };

  const resetData = () => {
    setData([]);
    setColumns([]);
    setSelectedRegions(new Set());
    setSearchTerm('');
  };

  // Helper to extract a short region name (e.g. "진접읍" from full address)
  const getShortRegion = (fullRegion: string) => {
    if (!fullRegion) return '기타';
    const parts = String(fullRegion).split(' ');
    // Try to find parts containing 읍, 면, 동
    const eumMyunDong = parts.find(p => p.endsWith('읍') || p.endsWith('면') || p.endsWith('동'));
    if (eumMyunDong) return eumMyunDong;
    // Fallback to second part if it's a typical Gyeonggi-do Namyangju-si XXX
    if (parts.length > 2 && parts[1] === '남양주시') return parts[2] || '기타';
    return parts[0] || '기타';
  };

  const extractNumber = (val: any) => {
    if (typeof val === 'number') return val;
    if (typeof val === 'string') {
      const num = parseFloat(val.replace(/,/g, '').replace(/[^0-9.]/g, ''));
      return isNaN(num) ? 0 : num;
    }
    return 0;
  };

  const processedData = useMemo(() => {
    return data.map((item, idx) => {
      const rawLocation = item[locationCol] || '';
      let region = regionCol && item[regionCol] ? String(item[regionCol]).trim() : getShortRegion(rawLocation);
      
      if (
        region === '다신지금로123번안길, 다신지금로145번안길' || 
        region === '다산지금로123번안길, 다산지금로145번안길' || 
        region.includes('다신지금로123번안길') || 
        region.includes('다산지금로123번안길')
      ) {
        region = '양정동';
      }

      return {
        ...item,
        _id: idx,
        _name: item[nameCol] || '-',
        _rawLocation: rawLocation,
        _region: region,
        _area: extractNumber(item[areaCol]),
        _url: urlCol ? item[urlCol] : undefined
      };
    });
  }, [data, nameCol, locationCol, areaCol, urlCol, regionCol]);

  const allRegions = useMemo(() => {
    const regions = new Set<string>();
    processedData.forEach(d => regions.add(d._region));
    return Array.from(regions).sort();
  }, [processedData]);

  const toggleRegion = (region: string) => {
    const newPaths = new Set(selectedRegions);
    if (newPaths.has(region)) newPaths.delete(region);
    else newPaths.add(region);
    setSelectedRegions(newPaths);
  };

  const filteredData = useMemo(() => {
    return processedData.filter(item => {
      const matchesSearch = item._name.toLowerCase().includes(searchTerm.toLowerCase()) || 
                            item._rawLocation.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesRegion = selectedRegions.size === 0 || selectedRegions.has(item._region);
      return matchesSearch && matchesRegion;
    });
  }, [processedData, searchTerm, selectedRegions]);

  const regionStats = useMemo(() => {
    const stats: Record<string, { count: number, totalArea: number }> = {};
    filteredData.forEach(item => {
      if (!stats[item._region]) stats[item._region] = { count: 0, totalArea: 0 };
      stats[item._region].count += 1;
      stats[item._region].totalArea += item._area;
    });
    return Object.entries(stats)
      .map(([name, data]) => ({ name, ...data }))
      .sort((a, b) => b.count - a.count);
  }, [filteredData]);

  const totalFacilities = filteredData.length;
  const totalArea = filteredData.reduce((acc, curr) => acc + curr._area, 0);

  if (data.length === 0) {
    return (
      <div className="min-h-screen bg-slate-900 text-slate-100 flex flex-col items-center justify-center p-6 font-sans">
        <div className="max-w-2xl w-full bg-slate-800/50 border border-slate-700 rounded-3xl shadow-xl overflow-hidden">
          <div className="border-b border-slate-700 p-8 text-center">
            <Building2 className="w-16 h-16 mx-auto mb-4 text-blue-400 opacity-80" />
            <h1 className="text-3xl font-bold tracking-tight">남양주시 물류창고 대시보드</h1>
            <p className="mt-2 text-slate-400 font-medium">관리하시는 엑셀 파일을 업로드하여 데이터를 시각화하세요</p>
          </div>
          
          <div className="p-8">
            <div 
              onDragOver={(e) => e.preventDefault()}
              onDrop={handleDrop}
              className="border border-dashed border-slate-600 rounded-2xl p-12 text-center hover:bg-slate-800 transition-colors cursor-pointer relative group"
            >
              <input 
                type="file" 
                accept=".xlsx, .xls, .csv" 
                onChange={handleFileUpload} 
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
              />
              <UploadCloud className="w-12 h-12 text-blue-400 mx-auto mb-4 group-hover:scale-110 transition-transform" />
              <p className="text-lg font-semibold text-slate-200">여기로 파일을 드래그하거나 클릭하여 업로드</p>
              <p className="text-sm text-slate-500 mt-2">지원 형식: .xlsx, .xls, .csv</p>
            </div>
            
            <div className="mt-8 bg-slate-800/80 rounded-xl p-5 border border-slate-700">
              <h3 className="text-sm font-bold text-slate-300 flex items-center mb-2">
                <FileSpreadsheet className="w-4 h-4 mr-2" /> 
                권장되는 데이터 구조
              </h3>
              <p className="text-sm text-slate-400 leading-relaxed">
                파일에 <span className="font-semibold px-1.5 py-0.5 rounded bg-slate-700 text-slate-200 border border-slate-600">상호명</span>, 
                <span className="font-semibold px-1.5 py-0.5 rounded bg-slate-700 text-slate-200 border border-slate-600 ml-1">소재지(주소)</span>, 
                <span className="font-semibold px-1.5 py-0.5 rounded bg-slate-700 text-slate-200 border border-slate-600 ml-1">면적</span>
                과 관련된 의미의 열 이름이 포함되어 있으면 대시보드가 자동으로 인식합니다. (예: "창고명", "읍면동", "연면적" 등 동일)
              </p>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 flex flex-col font-sans overflow-hidden">
      {/* Header */}
      <header className="px-6 py-5 flex items-center justify-between z-20 shrink-0">
        <div className="flex items-center space-x-3">
          <div className="w-10 h-10 rounded-full bg-slate-800 border border-slate-700 flex items-center justify-center text-blue-400">
            <Building2 className="w-5 h-5 text-blue-400" />
          </div>
          <div>
            <h1 className="text-2xl font-bold tracking-tight">NYJ WAREHOUSE <span className="text-blue-400">STATUS</span></h1>
            <p className="text-slate-400 text-xs">남양주시 물류창고 대시보드</p>
          </div>
        </div>
        
        <div className="flex flex-col md:flex-row md:items-center space-y-2 md:space-y-0 md:space-x-6">
           <div className="text-right hidden md:block">
              <p className="text-[10px] text-slate-500 uppercase font-semibold">총 로드된 데이터</p>
              <p className="text-sm font-mono text-emerald-400">{processedData.length.toLocaleString()} 건</p>
           </div>
           <button 
              onClick={resetData}
              className="flex items-center text-xs px-4 py-2 text-slate-300 bg-slate-800 border border-slate-700 hover:bg-slate-700 hover:text-white rounded-full font-bold transition-all"
           >
              <Trash2 className="w-3.5 h-3.5 mr-1.5" />
              새로운 파일 로드
           </button>
        </div>
      </header>

      <main className="flex-1 w-full p-6 pt-0 flex flex-col lg:flex-row gap-4 h-[calc(100vh-88px)]">
        
        {/* Sidebar Filters */}
        <aside className="w-full lg:w-[320px] flex-shrink-0 flex flex-col gap-4 overflow-y-auto custom-scrollbar pr-2">
          <div className="bg-slate-800/50 border border-slate-700 rounded-3xl p-6 shrink-0">
            <div className="flex items-center mb-4">
              <Search className="w-4 h-4 text-slate-500 mr-2" />
              <h2 className="text-sm font-bold uppercase text-slate-500 tracking-wider">검색</h2>
            </div>
            <input 
              type="text" 
              placeholder="상호 또는 주소 검색..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full px-4 py-2.5 bg-slate-900 border border-slate-700 rounded-xl text-sm focus:outline-none focus:border-slate-500 text-slate-200 placeholder-slate-600 transition-colors"
            />
          </div>

          <div className="bg-slate-800/50 border border-slate-700 rounded-3xl p-6 flex-1 min-h-0 flex flex-col">
            <div className="flex flex-col mb-4 shrink-0">
              <div className="flex items-center justify-between mb-1">
                <div className="flex items-center">
                  <Filter className="w-4 h-4 text-slate-500 mr-2" />
                  <h2 className="text-sm font-bold uppercase text-slate-500 tracking-wider">지역 필터</h2>
                </div>
                {selectedRegions.size > 0 && (
                  <button 
                    onClick={() => setSelectedRegions(new Set())}
                    className="text-[10px] uppercase font-bold text-blue-400 hover:text-blue-300 bg-blue-500/10 px-2 py-0.5 rounded-full"
                  >
                    초기화
                  </button>
                )}
              </div>
              <p className="text-xs text-slate-500">다중 선택 가능</p>
            </div>
            <div className="overflow-y-auto pr-2 space-y-1 custom-scrollbar flex-1">
              {allRegions.map(region => (
                <label key={region} className="flex items-center group cursor-pointer p-1.5 -mx-1.5 rounded-lg hover:bg-slate-700/50 transition-colors">
                  <input 
                    type="checkbox" 
                    checked={selectedRegions.has(region)}
                    onChange={() => toggleRegion(region)}
                    className="rounded border-slate-600 bg-slate-800 text-blue-500 focus:ring-blue-500 focus:ring-offset-slate-900 mr-3 w-4 h-4"
                  />
                  <span className={`text-sm ${selectedRegions.has(region) ? 'text-white font-medium' : 'text-slate-400 group-hover:text-slate-200'}`}>{region}</span>
                </label>
              ))}
            </div>
          </div>
          
          <div className="bg-slate-800/50 border border-slate-700 rounded-3xl p-6 shrink-0">
            <h2 className="text-sm font-bold uppercase text-slate-500 tracking-wider mb-4">데이터 자동 매핑</h2>
            <div className="space-y-3 text-sm text-slate-400">
              <div className="flex justify-between items-center border-b border-slate-700/50 pb-2">
                <span className="font-medium text-xs uppercase">지역구분</span>
                <select 
                   value={regionCol} 
                   onChange={e => setRegionCol(e.target.value)} 
                   className="bg-slate-900 border border-slate-700 rounded-md px-2 py-1 font-semibold focus:outline-none max-w-[140px] truncate text-slate-300 text-xs"
                   title={regionCol}
                >
                  <option value="">-- 자동생성 --</option>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div className="flex justify-between items-center border-b border-slate-700/50 pb-2">
                <span className="font-medium text-xs uppercase">상호</span>
                <select 
                   value={nameCol} 
                   onChange={e => setNameCol(e.target.value)} 
                   className="bg-slate-900 border border-slate-700 rounded-md px-2 py-1 font-semibold focus:outline-none max-w-[140px] truncate text-slate-300 text-xs"
                   title={nameCol}
                >
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div className="flex justify-between items-center border-b border-slate-700/50 pb-2">
                <span className="font-medium text-xs uppercase">주소</span>
                <select 
                   value={locationCol} 
                   onChange={e => setLocationCol(e.target.value)} 
                   className="bg-slate-900 border border-slate-700 rounded-md px-2 py-1 font-semibold focus:outline-none max-w-[140px] truncate text-slate-300 text-xs"
                   title={locationCol}
                >
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div className="flex justify-between items-center border-b border-slate-700/50 pb-2">
                <span className="font-medium text-xs uppercase">면적</span>
                <select 
                   value={areaCol} 
                   onChange={e => setAreaCol(e.target.value)} 
                   className="bg-slate-900 border border-slate-700 rounded-md px-2 py-1 font-semibold focus:outline-none max-w-[140px] truncate text-slate-300 text-xs"
                   title={areaCol}
                >
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div className="flex justify-between items-center">
                <span className="font-medium text-xs uppercase">링크</span>
                <select 
                   value={urlCol} 
                   onChange={e => setUrlCol(e.target.value)} 
                   className="bg-slate-900 border border-slate-700 rounded-md px-2 py-1 font-semibold focus:outline-none max-w-[140px] truncate text-slate-300 text-xs"
                   title={urlCol}
                >
                  <option value="">-- 자동 --</option>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
            </div>
          </div>
        </aside>

        {/* Main Content */}
        <div className="flex-1 flex flex-col gap-4 min-w-0 h-full overflow-hidden">
          
          {/* Summary Cards */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 shrink-0">
            <div className="bg-slate-800/50 border border-slate-700 rounded-3xl p-6 flex flex-col justify-between">
              <div className="text-slate-500 text-xs font-bold uppercase tracking-wider mb-2">필터링된 창고 수</div>
              <div>
                <div className="text-3xl lg:text-4xl font-bold text-white tracking-tight">
                  {totalFacilities.toLocaleString()} <span className="text-sm text-slate-500 font-normal ml-1">개소</span>
                </div>
              </div>
            </div>
            
            <div className="bg-slate-800/50 border border-slate-700 rounded-3xl p-6 flex flex-col justify-between">
              <div className="text-slate-500 text-xs font-bold uppercase tracking-wider mb-2">총 면적 합계</div>
              <div>
                <div className="text-3xl lg:text-4xl font-bold text-white tracking-tight">
                   {totalArea >= 1000000 
                     ? (totalArea / 1000000).toLocaleString(undefined, {maximumFractionDigits: 2}) + 'M'
                     : totalArea >= 1000 
                       ? (totalArea / 1000).toLocaleString(undefined, {maximumFractionDigits: 1}) + 'K'
                       : totalArea.toLocaleString(undefined, {maximumFractionDigits: 0})
                   }
                   <span className="text-sm text-slate-500 font-normal ml-1">㎡</span>
                </div>
                <div className="w-full bg-slate-700 h-1.5 rounded-full mt-3 overflow-hidden">
                  <div className="bg-blue-500 h-full w-[80%]"></div>
                </div>
              </div>
            </div>

            <div className="bg-slate-800/50 border border-slate-700 rounded-3xl p-6 flex flex-col justify-between">
               <div className="text-slate-500 text-xs font-bold uppercase tracking-wider mb-2">해당 지역 수</div>
              <div>
                <div className="text-3xl lg:text-4xl font-bold text-white tracking-tight">
                   {new Set(filteredData.map(d => d._region)).size} <span className="text-sm text-slate-500 font-normal ml-1">지역</span>
                </div>
              </div>
            </div>
          </div>

          {/* Chart and Table Row */}
          <div className="flex-1 grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-5 gap-4 min-h-0">
            {/* Rank List Section */}
            <div className="xl:col-span-2 bg-slate-800/50 border border-slate-700 rounded-3xl p-6 flex flex-col min-h-0">
              <div className="flex items-center space-x-2 mb-4 shrink-0">
                <MapPin className="w-4 h-4 text-slate-500" />
                <h3 className="text-sm font-bold uppercase text-slate-500 tracking-wider">지역별 창고 수 (상위 10개)</h3>
              </div>
              <div className="flex-1 w-full min-h-0 overflow-y-auto custom-scrollbar pr-2">
                {regionStats.length > 0 ? (
                  <div className="space-y-3">
                    {regionStats.slice(0, 10).map((entry, index) => (
                      <div 
                        key={index} 
                        className={`flex justify-between items-center p-3 rounded-xl border transition-colors cursor-pointer ${selectedRegions.has(entry.name) ? 'bg-blue-500/20 border-blue-500/50' : 'bg-slate-700/20 border-slate-700/50 hover:bg-slate-700/40'}`}
                        onClick={() => toggleRegion(entry.name)}
                      >
                        <div className="flex items-center space-x-3">
                          <span className={`flex items-center justify-center w-6 h-6 rounded-full text-xs font-bold ${selectedRegions.has(entry.name) ? 'bg-blue-500 text-white' : 'bg-slate-800 border-slate-600 border text-slate-400'}`}>
                            {index + 1}
                          </span>
                          <span className={`text-sm font-medium ${selectedRegions.has(entry.name) ? 'text-white' : 'text-slate-200'}`}>{entry.name}</span>
                        </div>
                        <div className="text-right">
                          <div className={`font-mono font-bold text-lg ${selectedRegions.has(entry.name) ? 'text-blue-300' : 'text-blue-400'}`}>
                             {entry.count.toLocaleString()}
                             <span className="text-xs text-slate-500 font-sans ml-1">개소</span>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="w-full h-full flex flex-col items-center justify-center text-slate-500">
                    <MapPin className="w-8 h-8 mb-2 opacity-50" />
                    <p className="text-sm tracking-wide">데이터 없음</p>
                  </div>
                )}
              </div>
            </div>

            {/* Data Table */}
            <div className="xl:col-span-3 bg-slate-800/50 border border-slate-700 rounded-3xl p-6 overflow-hidden flex flex-col min-h-0">
              <div className="flex justify-between items-center mb-4 shrink-0">
                <div className="flex items-center space-x-2">
                  <TableProperties className="w-4 h-4 text-slate-500" />
                  <h3 className="text-sm font-bold uppercase text-slate-500 tracking-wider">상세 현황표</h3>
                </div>
                <span className="px-3 py-1 bg-slate-700 rounded text-[10px] font-bold tracking-widest text-emerald-400 border border-slate-600/50 shadow-inner uppercase">LIVE DATA</span>
              </div>
              
              <div className="overflow-auto flex-1 relative custom-scrollbar pr-2">
                <table className="w-full text-left border-collapse min-w-max">
                  <thead className="text-xs text-slate-500 border-b border-slate-700 uppercase sticky top-0 z-10 bg-slate-800/90 backdrop-blur">
                    <tr>
                      <th className="pb-3 px-4 font-semibold whitespace-nowrap">상호</th>
                      <th className="pb-3 px-4 font-semibold whitespace-nowrap">지역</th>
                      <th className="pb-3 px-4 font-semibold whitespace-nowrap">소재지</th>
                      <th className="pb-3 px-4 font-semibold text-right whitespace-nowrap">면적 (㎡)</th>
                    </tr>
                  </thead>
                  <tbody className="text-sm divide-y divide-slate-700/50">
                    {filteredData.length > 0 ? (
                      filteredData.map((row) => (
                        <tr key={row._id} className="hover:bg-slate-700/30 transition-colors group">
                          <td className="py-3 px-4 font-medium text-slate-200">
                             {row._url ? (
                               <a href={row._url} target="_blank" rel="noopener noreferrer" className="hover:text-blue-400 hover:underline transition-colors cursor-pointer">
                                 {row._name}
                               </a>
                             ) : (
                               row._name
                             )}
                          </td>
                          <td className="py-3 px-4">
                            <span className="inline-flex items-center px-2 py-0.5 rounded-full bg-slate-700/80 text-[11px] font-bold text-slate-300 border border-slate-600">
                              {row._region}
                            </span>
                          </td>
                          <td className="py-3 px-4 text-slate-400 text-xs max-w-[200px] truncate group-hover:text-slate-300 transition-colors" title={row._rawLocation}>
                            {row._rawLocation}
                          </td>
                          <td className="py-3 px-4 text-slate-300 font-mono text-right tabular-nums">
                            {row._area ? row._area.toLocaleString(undefined, {maximumFractionDigits: 1}) : '-'}
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={4} className="py-12 text-center text-slate-500 text-sm">
                          조건에 맞는 데이터가 없습니다.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}

