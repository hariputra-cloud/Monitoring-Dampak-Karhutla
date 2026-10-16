<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Monitoring Dampak Gangguan</title>
    
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    
    <!-- React & ReactDOM CDN -->
    <script crossorigin src="https://unpkg.com/react@18/umd/react.production.min.js"></script>
    <script crossorigin src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></script>
    
    <!-- Babel standalone untuk transpile JSX di browser -->
    <script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
    
    <!-- Recharts CDN & Prop-types dependency -->
    <script src="https://unpkg.com/prop-types@15.8.1/prop-types.min.js"></script>
    <script src="https://unpkg.com/recharts@2.12.7/umd/Recharts.js"></script>
    
    <!-- SheetJS CDN (Excel Parser) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    
    <!-- Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>
</head>
<body class="bg-[#0b1121]">

    <div id="root"></div>

    <script type="text/babel">
        const { useState, useEffect, useMemo, useRef } = React;
        const { 
            BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, 
            LineChart, Line, PieChart, Pie, Cell, Sector, Area, AreaChart, LabelList 
        } = window.Recharts;

        // Mock Component Icon sederhana untuk menggantikan Lucide Icons dalam lingkungan React browser
        const Icon = ({ name, className = "w-5 h-5" }) => {
            return <i data-lucide={name} className={className}></i>;
        };

        const COLORS = ['#38bdf8', '#34d399', '#fbbf24', '#f87171', '#a78bfa', '#2dd4bf', '#fb923c', '#a3e635', '#22d3ee'];

        function parseDateString(val) {
            if (!val) return new Date();
            if (val instanceof Date) return new Date(val.getFullYear(), val.getMonth(), val.getDate());
            if (typeof val === 'number') {
                const d = new Date(Math.round((val - 25569) * 86400 * 1000));
                return new Date(d.getFullYear(), d.getMonth(), d.getDate());
            }
            let strVal = String(val).trim().split(' ')[0].split('T')[0];
            const parts = strVal.split(/[-/]/);
            if (parts.length === 3) {
                if (parts[2].length === 4) return new Date(parts[2], parseInt(parts[1]) - 1, parts[0]);
                if (parts[0].length === 4) return new Date(parts[0], parseInt(parts[1]) - 1, parts[2]);
            }
            const d = new Date(strVal);
            if (isNaN(d.getTime())) return new Date();
            return new Date(d.getFullYear(), d.getMonth(), d.getDate());
        }

        const CustomTooltip = ({ active, payload, label }) => {
            if (active && payload && payload.length) {
                return (
                    <div className="bg-slate-800/95 backdrop-blur-sm p-4 rounded-xl shadow-2xl border border-slate-700">
                        <p className="font-semibold text-slate-200 mb-3 border-b border-slate-700 pb-2">{label}</p>
                        {payload.map((entry, index) => (
                            <p key={`item-${index}`} className="text-sm flex items-center justify-between gap-4 mb-1" style={{ color: entry.color || entry.fill }}>
                                <span className="flex items-center gap-2">
                                    <span className="w-2.5 h-2.5 rounded-full inline-block shadow-[0_0_8px_currentColor]" style={{ backgroundColor: entry.color || entry.fill }}></span>
                                    {entry.name}
                                </span>
                                <span className="font-bold text-slate-100">{entry.value}</span>
                            </p>
                        ))}
                    </div>
                );
            }
            return null;
        };

        function App() {
            const [data, setData] = useState([]);
            const [loading, setLoading] = useState(false);
            const [error, setError] = useState(null);
            const [fileName, setFileName] = useState(null);
            
            const [selectedAirport, setSelectedAirport] = useState('All');
            const [timeView, setTimeView] = useState('daily');
            const [startDate, setStartDate] = useState('');
            const [endDate, setEndDate] = useState('');
            
            const fileInputRef = useRef(null);

            useEffect(() => {
                if (window.lucide) {
                    window.lucide.createIcons();
                }
            });

            const handleFileUpload = (e) => {
                const file = e.target.files[0];
                if (!file) return;
                
                setLoading(true);
                setError(null);

                const reader = new FileReader();
                reader.onload = (evt) => {
                    try {
                        const arrayBuffer = evt.target.result;
                        const dataUint8 = new Uint8Array(arrayBuffer);
                        const workbook = window.XLSX.read(dataUint8, { type: 'array', cellDates: true });
                        
                        const firstSheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheetName];
                        
                        const rawData = window.XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        let headerRowIndex = 0;
                        let maxScore = -1;
                        
                        for (let i = 0; i < Math.min(20, rawData.length); i++) {
                            const row = rawData[i];
                            if (!row || !Array.isArray(row)) continue;
                            
                            let score = 0;
                            const rowStr = row.map(cell => String(cell).toLowerCase()).join(' ');
                            
                            if (rowStr.match(/tanggal|date|waktu|hari|tgl/)) score += 3;
                            if (rowStr.match(/bandara|airport|stasiun|station|lokasi|branch|cabang/)) score += 3;
                            if (rowStr.match(/jenis|kategori|type|gangguan|irreg|masalah/)) score += 3;
                            score += row.filter(cell => typeof cell === 'string' && cell.trim() !== '').length * 0.5;
                            
                            if (score > maxScore) {
                                maxScore = score;
                                headerRowIndex = i;
                            }
                        }

                        const jsonData = window.XLSX.utils.sheet_to_json(worksheet, { range: headerRowIndex, defval: '' });
                        if (!jsonData || jsonData.length === 0) throw new Error("Data Excel kosong atau format tabel tidak ditemukan.");

                        const headers = Object.keys(jsonData[0]);
                        const findCol = (regex) => headers.find(h => h && String(h).toLowerCase().match(regex));
                        
                        let dateCol = findCol(/tanggal|date|waktu|hari|tgl/i);
                        let airportCol = findCol(/bandara|airport|stasiun|station|lokasi|branch|cabang/i);
                        let typeCol = findCol(/jenis|kategori|type|gangguan|irreg|masalah|deskripsi|keterangan/i);

                        const validHeaders = headers.filter(h => h && !h.startsWith('__EMPTY'));
                        if (!dateCol) dateCol = validHeaders[0] || headers[0];
                        if (!airportCol) airportCol = validHeaders[1] || headers[1];
                        if (!typeCol) typeCol = validHeaders[2] || headers[2];

                        const processedData = jsonData.map(row => {
                            const rawDate = row[dateCol];
                            const parsedDate = parseDateString(rawDate);
                            
                            const yyyy = parsedDate.getFullYear();
                            const mm = String(parsedDate.getMonth() + 1).padStart(2, '0');
                            const dd = String(parsedDate.getDate()).padStart(2, '0');
                            
                            return {
                                date: parsedDate,
                                dateStr: `${yyyy}-${mm}-${dd}`,
                                monthStr: `${yyyy}-${mm}`,
                                airport: row[airportCol] ? String(row[airportCol]).trim() : 'Tanpa Bandara',
                                type: row[typeCol] ? String(row[typeCol]).trim() : 'Lain-lain',
                                rawDateStr: String(rawDate).toLowerCase()
                            };
                        }).filter(d => {
                            const isHeader = String(d.type).toLowerCase() === String(typeCol).toLowerCase() || d.rawDateStr.includes('tanggal') || d.rawDateStr.includes('date');
                            const isValidDate = !isNaN(d.date.getTime());
                            const isExcludedType = String(d.type).trim().toLowerCase() === 'laporan kejadian (incident/accident)';
                            return !isHeader && isValidDate && d.airport !== 'undefined' && d.airport !== '' && !isExcludedType;
                        });

                        if (processedData.length === 0) throw new Error("Gagal mengekstrak data baris.");

                        const dates = processedData.map(d => d.date.getTime());
                        const formatYMD = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
                        
                        setStartDate(formatYMD(new Date(Math.min(...dates))));
                        setEndDate(formatYMD(new Date(Math.max(...dates))));
                        setData(processedData);
                        setFileName(file.name);
                        setLoading(false);
                    } catch (err) {
                        setError(err.message || "Gagal membaca format file.");
                        setLoading(false);
                    }
                };
                reader.readAsArrayBuffer(file);
                e.target.value = null;
            };

            const { airports, totalByAirport, trendData, totalRecords, mostFrequentType, highestAirport } = useMemo(() => {
                if (!data.length) return { airports: [], totalByAirport: [], trendData: [], totalRecords: 0, mostFrequentType: '-', highestAirport: '-' };

                const periodData = data.filter(d => (!startDate || !endDate) || (d.dateStr >= startDate && d.dateStr <= endDate));
                const uniqueAirports = Array.from(new Set(periodData.map(d => d.airport))).sort();
                const filteredData = selectedAirport === 'All' ? periodData : periodData.filter(d => d.airport === selectedAirport);

                const airportCounts = {};
                const globalTypeCounts = {};
                periodData.forEach(d => {
                    airportCounts[d.airport] = (airportCounts[d.airport] || 0) + 1;
                    globalTypeCounts[d.type] = (globalTypeCounts[d.type] || 0) + 1;
                });

                const overallAirportsData = Object.entries(airportCounts).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
                const overallTypesData = Object.entries(globalTypeCounts).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);

                const trendMap = {};
                filteredData.forEach(d => {
                    const key = timeView === 'daily' ? d.dateStr : d.monthStr;
                    if (!trendMap[key]) trendMap[key] = { name: key, total: 0 };
                    trendMap[key].total += 1;
                });

                return {
                    airports: uniqueAirports,
                    totalByAirport: overallAirportsData,
                    trendData: Object.values(trendMap).sort((a, b) => a.name.localeCompare(b.name)),
                    totalRecords: periodData.length,
                    mostFrequentType: overallTypesData[0]?.name || '-',
                    highestAirport: overallAirportsData[0] ? `${overallAirportsData[0].name} (${overallAirportsData[0].value})` : '-'
                };
            }, [data, selectedAirport, timeView, startDate, endDate]);

            return (
                <div className="min-h-screen bg-[#0b1121] text-slate-300 font-sans p-4 md:p-6 selection:bg-cyan-900 selection:text-cyan-100">
                    <div className="max-w-[1400px] mx-auto space-y-6">
                        
                        {/* Header */}
                        <div className="bg-[#111827]/80 backdrop-blur-md rounded-2xl border border-slate-800/80 p-5 flex flex-col xl:flex-row justify-between items-start xl:items-center gap-6 shadow-2xl">
                            <div className="flex items-center gap-4">
                                <div className="w-14 h-14 bg-gradient-to-br from-cyan-500 to-blue-600 rounded-xl flex items-center justify-center shrink-0">
                                    <Icon name="activity" className="w-8 h-8 text-white" />
                                </div>
                                <div>
                                    <h1 className="text-2xl md:text-3xl font-black text-transparent bg-clip-text bg-gradient-to-r from-cyan-400 to-blue-400">
                                        MONITORING DAMPAK GANGGUAN
                                    </h1>
                                    <div className="flex items-center gap-2 mt-1.5">
                                        <span className="flex items-center gap-1.5 text-xs font-medium px-2.5 py-1 rounded-md bg-slate-800 text-slate-400 border border-slate-700">
                                            <Icon name="calendar" className="w-3.5 h-3.5" /> Data Operasional
                                        </span>
                                        {fileName && (
                                            <span className="flex items-center gap-1.5 text-xs font-medium px-2.5 py-1 rounded-md bg-emerald-900/30 text-emerald-400 border border-emerald-800/50">
                                                <Icon name="check-circle-2" className="w-3.5 h-3.5" /> {fileName}
                                            </span>
                                        )}
                                    </div>
                                </div>
                            </div>
                            
                            <div className="flex flex-col sm:flex-row gap-3 w-full xl:w-auto">
                                <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} ref={fileInputRef} className="hidden" />
                                <button 
                                    onClick={() => fileInputRef.current?.click()}
                                    disabled={loading}
                                    className="flex items-center justify-center gap-2 px-5 py-3 rounded-xl font-semibold bg-blue-600/90 hover:bg-blue-500 text-white border border-blue-500/50 shadow-lg"
                                >
                                    {loading ? <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white"></div> : <Icon name="upload-cloud" />}
                                    <span>Upload Excel Data</span>
                                </button>
                                
                                {data.length > 0 && (
                                    <div className="flex items-center gap-2 bg-[#1e293b] px-4 py-3 rounded-xl border border-slate-700">
                                        <Icon name="map-pin" className="w-5 h-5 text-cyan-400" />
                                        <select 
                                            value={selectedAirport} 
                                            onChange={(e) => setSelectedAirport(e.target.value)}
                                            className="bg-transparent border-none text-slate-200 font-medium outline-none cursor-pointer"
                                        >
                                            <option value="All">Semua Bandara (Nasional)</option>
                                            {airports.map(apt => (
                                                <option key={apt} value={apt} className="bg-slate-800">{apt}</option>
                                            ))}
                                        </select>
                                    </div>
                                )}
                            </div>
                        </div>

                        {/* Filter Bar */}
                        {data.length > 0 && (
                            <div className="bg-[#111827]/80 rounded-xl border border-slate-800 p-4 flex flex-col md:flex-row items-center gap-4">
                                <div className="flex items-center gap-3 text-slate-300">
                                    <Icon name="calendar" className="w-5 h-5 text-cyan-400" />
                                    <span className="font-semibold text-sm uppercase">Filter Periode:</span>
                                </div>
                                <div className="flex items-center gap-2 bg-[#1e293b] rounded-lg p-1.5 border border-slate-700">
                                    <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} className="bg-slate-800 text-slate-200 px-3 py-1.5 rounded-md text-sm outline-none" />
                                    <span className="text-slate-500 font-bold px-1">s/d</span>
                                    <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)} className="bg-slate-800 text-slate-200 px-3 py-1.5 rounded-md text-sm outline-none" />
                                </div>
                            </div>
                        )}

                        {/* Content Area */}
                        {data.length === 0 ? (
                            <div className="bg-[#111827]/50 rounded-2xl border border-dashed border-slate-700 h-[50vh] flex flex-col justify-center items-center text-slate-500 gap-4 text-center">
                                <Icon name="file-spreadsheet" className="w-16 h-16 text-slate-600" />
                                <h3 className="text-xl font-bold text-slate-300">Sistem Siaga</h3>
                                <p className="max-w-md text-slate-500">Unggah file Excel laporan gangguan operasional untuk menampilkan grafik analitik.</p>
                                <button onClick={() => fileInputRef.current?.click()} className="bg-cyan-600/20 text-cyan-400 border border-cyan-500/30 px-5 py-2.5 rounded-xl font-semibold flex items-center gap-2">
                                    <Icon name="upload-cloud" /> Upload File
                                </button>
                            </div>
                        ) : (
                            <div className="space-y-6">
                                {/* KPI Cards */}
                                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                                    {[
                                        { title: 'Total Keseluruhan Data', val: totalRecords, icon: 'activity', color: 'text-cyan-400' },
                                        { title: 'Total Bandara Terdampak', val: airports.length, icon: 'map-pin', color: 'text-emerald-400' },
                                        { title: 'Frekuensi Tertinggi', val: highestAirport, icon: 'bar-chart-3', color: 'text-orange-400' },
                                        { title: 'Dominasi Gangguan', val: mostFrequentType, icon: 'shield-alert', color: 'text-pink-400' }
                                    ].map((card, i) => (
                                        <div key={i} className="bg-[#111827]/80 rounded-2xl p-5 border border-slate-800 flex items-center gap-4">
                                            <div className="p-3 bg-slate-800 rounded-xl">
                                                <Icon name={card.icon} className={`w-6 h-6 ${card.color}`} />
                                            </div>
                                            <div>
                                                <p className="text-xs font-semibold text-slate-400 uppercase">{card.title}</p>
                                                <p className="text-xl font-bold text-white mt-1">{card.val}</p>
                                            </div>
                                        </div>
                                    ))}
                                </div>

                                {/* Trend Chart */}
                                <div className="bg-[#111827]/80 rounded-2xl p-6 border border-slate-800 h-[400px] flex flex-col">
                                    <div className="flex justify-between items-center mb-4">
                                        <h2 className="text-lg font-bold text-slate-100 flex items-center gap-2">
                                            <Icon name="trending-up" className="text-blue-400" />
                                            TREND GANGGUAN HARIAN & BULANAN
                                        </h2>
                                        <div className="flex bg-[#1e293b] rounded-lg p-1 border border-slate-700 text-sm">
                                            <button onClick={() => setTimeView('daily')} className={`px-3 py-1 rounded-md ${timeView === 'daily' ? 'bg-blue-600 text-white' : 'text-slate-400'}`}>Harian</button>
                                            <button onClick={() => setTimeView('monthly')} className={`px-3 py-1 rounded-md ${timeView === 'monthly' ? 'bg-blue-600 text-white' : 'text-slate-400'}`}>Bulanan</button>
                                        </div>
                                    </div>
                                    <div className="flex-1 w-full min-h-0">
                                        <ResponsiveContainer width="100%" height="100%">
                                            <AreaChart data={trendData}>
                                                <CartesianGrid strokeDasharray="3 3" stroke="#334155" opacity={0.5} />
                                                <XAxis dataKey="name" tick={{ fontSize: 11, fill: '#94a3b8' }} />
                                                <YAxis tick={{ fontSize: 11, fill: '#94a3b8' }} />
                                                <Tooltip content={<CustomTooltip />} />
                                                <Area type="monotone" dataKey="total" name="Total Insiden" stroke="#38bdf8" fill="#38bdf8" fillOpacity={0.2} strokeWidth={3}>
                                                    <LabelList dataKey="total" position="top" fill="#94a3b8" fontSize={11} />
                                                </Area>
                                            </AreaChart>
                                        </ResponsiveContainer>
                                    </div>
                                </div>

                                {/* Bar Chart */}
                                <div className="bg-[#111827]/80 rounded-2xl p-6 border border-slate-800 h-[400px] flex flex-col">
                                    <h2 className="text-lg font-bold text-slate-100 mb-4 flex items-center gap-2">
                                        <Icon name="bar-chart-3" className="text-emerald-400" />
                                        JUMLAH TOTAL GANGGUAN PER BANDARA
                                    </h2>
                                    <div className="flex-1 w-full min-h-0">
                                        <ResponsiveContainer width="100%" height="100%">
                                            <BarChart data={totalByAirport}>
                                                <CartesianGrid strokeDasharray="3 3" stroke="#334155" opacity={0.5} />
                                                <XAxis dataKey="name" tick={{ fontSize: 11, fill: '#94a3b8' }} />
                                                <YAxis tick={{ fontSize: 11, fill: '#94a3b8' }} />
                                                <Tooltip content={<CustomTooltip />} />
                                                <Bar dataKey="value" name="Jumlah Gangguan" fill="#34d399" radius={[4, 4, 0, 0]} />
                                            </BarChart>
                                        </ResponsiveContainer>
                                    </div>
                                </div>
                            </div>
                        )}
                    </div>
                </div>
            );
        }

        const root = ReactDOM.createRoot(document.getElementById('root'));
        root.render(<App />);
    </script>
</body>
</html>