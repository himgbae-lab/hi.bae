/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
  Legend,
  RadarChart,
  PolarGrid,
  PolarAngleAxis,
  PolarRadiusAxis,
  Radar
} from 'recharts';
import { Upload, FileSpreadsheet, BarChart3, PieChart as PieChartIcon, Table as TableIcon, X, ChevronRight, TrendingUp, TrendingDown, Minus, FileText, Sparkles, Printer, Download } from 'lucide-react';
import { cn } from './lib/utils';
import { GoogleGenAI } from "@google/genai";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

interface SheetData {
  name: string;
  headers: string[];
  rows: any[];
  categories: {
    effect: string | null;
    curriculum: string | null;
    operation: string | null;
    week: string | null;
    course: string | null;
  };
  summary: {
    totalRows: number;
    totalCols: number;
  };
}

interface DashboardData {
  sheets: SheetData[];
  activeSheetIndex: number;
}

const COLORS = ['#3B82F6', '#10B981', '#F59E0B', '#6366F1', '#EC4899'];

const WEEK_BLACKLIST = ['교과편성', '교육운영', '교육효과', '구분', 'N/A', 'n/a', 'nan', 'NaN'];

export default function App() {
  const [data, setData] = useState<DashboardData | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [activeTab, setActiveTab] = useState<'overview' | 'table' | 'report'>('overview');
  const [selectedWeek, setSelectedWeek] = useState<string | null>(null);
  const [isGeneratingReport, setIsGeneratingReport] = useState(false);
  const [executiveReport, setExecutiveReport] = useState<{ week1: string, week2: string, total: string } | null>(null);

  const parseNumeric = (val: any): number | null => {
    if (val === null || val === undefined || val === "") return null;
    if (typeof val === 'number') return val;
    
    let str = String(val).trim();

    const likertMap: { [key: string]: number } = {
      '매우 만족': 5, '매우만족': 5, '매우 우수': 5, '매우우수': 5, '매우 그렇다': 5, '매우그렇다': 5,
      '만족': 4, '우수': 4, '그렇다': 4,
      '보통': 3, '그저 그렇다': 3, '그저그렇다': 3,
      '불만족': 2, '미흡': 2, '그렇지 않다': 2, '그렇지않다': 2,
      '매우 불만족': 1, '매우불만족': 1, '매우 미흡': 1, '매우미흡': 1, '매우 그렇지 않다': 1, '매우그렇지않다': 1
    };
    
    for (const [key, num] of Object.entries(likertMap)) {
      if (str.includes(key)) return num;
    }
    
    if (str.includes(',') && !str.includes('.')) {
      str = str.replace(',', '.');
    }
    
    const cleaned = str.replace(/[^0-9.]/g, '');
    const parsed = parseFloat(cleaned);
    
    if (isNaN(parsed)) return null;
    
    if (parsed > 10 && parsed <= 50) return parsed / 10;
    
    return parsed;
  };

  const processFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const bstr = e.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary', cellDates: true, cellNF: true, cellText: true });
      
      const sheets: SheetData[] = wb.SheetNames.map(wsname => {
        const ws = wb.Sheets[wsname];
        const rawData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any[][];

        if (rawData.length === 0) return null;

        let headerRowIndex = 0;
        for (let i = 0; i < Math.min(rawData.length, 20); i++) {
          const nonEmptyCount = rawData[i].filter(cell => String(cell).trim() !== "").length;
          if (nonEmptyCount >= 2) {
            headerRowIndex = i;
            break;
          }
        }

        const headers = rawData[headerRowIndex].map((h, idx) => {
          const val = String(h).trim();
          return val !== "" ? val : `Column_${idx + 1}`;
        });

        const findCol = (keywords: string[]) => {
          const exact = headers.find(h => keywords.some(k => h === k));
          if (exact) return exact;
          const satisfactionMatch = headers.find(h => 
            keywords.some(k => h.includes(k)) && 
            (h.includes('만족') || h.includes('점수') || h.includes('평가') || h.includes('척도'))
          );
          if (satisfactionMatch) return satisfactionMatch;
          return headers.find(h => keywords.some(k => h.toLowerCase().includes(k.toLowerCase()))) || null;
        };

        const categories = {
          effect: findCol(['교육효과', '효과', '강사', '강의']),
          curriculum: findCol(['교과편성', '편성', '커리큘럼', '내용']),
          operation: findCol(['교육운영', '운영', '환경', '지원', '시설']),
          week: findCol(['주차', 'Week', '시기', '날짜']),
          course: findCol(['과정', 'Course', '프로그램', '반', '기수'])
        };

        const rows = rawData.slice(headerRowIndex + 1)
          .filter(row => row.some(cell => String(cell).trim() !== ""))
          .map(row => {
            const obj: any = {};
            headers.forEach((h, i) => {
              const val = row[i] === "" ? null : row[i];
              if ([categories.effect, categories.curriculum, categories.operation].includes(h)) {
                obj[h] = parseNumeric(val);
              } else {
                obj[h] = val;
              }
            });
            return obj;
          });

        return {
          name: wsname,
          headers,
          rows,
          categories,
          summary: {
            totalRows: rows.length,
            totalCols: headers.length
          }
        };
      }).filter((s): s is SheetData => s !== null);

      if (sheets.length > 0) {
        const firstSheet = sheets[0];
        const weekKey = firstSheet.categories.week || firstSheet.headers[0];
        const uniqueWeeks = Array.from(new Set(firstSheet.rows.map(r => String(r[weekKey] || '').trim())))
          .filter((w: any): w is string => !!w && !WEEK_BLACKLIST.includes(w))
          .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
        const defaultWeek = uniqueWeeks.find(w => w.includes('1')) || uniqueWeeks[0] || null;

        setData({
          sheets,
          activeSheetIndex: 0
        });
        setSelectedWeek(defaultWeek);
      }
    };
    reader.readAsBinaryString(file);
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.name.endsWith('.csv'))) {
      processFile(file);
    }
  }, []);

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const currentSheet = data?.sheets[data.activeSheetIndex];

  // Simple heuristic to find numeric columns for charts
  const numericCols = currentSheet?.headers.filter(h => 
    currentSheet.rows.some(row => typeof row[h] === 'number')
  ) || [];

  const stringCols = currentSheet?.headers.filter(h => 
    currentSheet.rows.some(row => typeof row[h] === 'string')
  ) || [];

  const chartData = currentSheet?.rows.slice(0, 10) || [];

  const weekKey = currentSheet?.categories.week || currentSheet?.headers[0] || '';
  const uniqueWeeks = Array.from(new Set(currentSheet?.rows.map(r => String(r[weekKey] || '').trim()) || []))
    .filter((w: any): w is string => !!w && !WEEK_BLACKLIST.includes(w))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  const filteredRows = selectedWeek 
    ? currentSheet?.rows.filter(r => String(r[weekKey] || '') === selectedWeek) || []
    : currentSheet?.rows || [];

  const [showDiagnostics, setShowDiagnostics] = useState(false);

  // Helper to calculate average for a specific category
  const getAverage = (rows: any[], categoryKey: string | null) => {
    if (!categoryKey) return 0;
    const vals = rows.map(r => r[categoryKey]).filter(v => typeof v === 'number');
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
  };

  // Helper to get distribution for Pie Chart
  const getDistribution = (rows: any[]) => {
    if (!currentSheet) return [];
    const keys = [currentSheet.categories.effect, currentSheet.categories.curriculum, currentSheet.categories.operation].filter(Boolean) as string[];
    const allVals = rows.flatMap(r => keys.map(k => r[k])).filter(v => typeof v === 'number');
    const counts: { [key: number]: number } = { 5: 0, 4: 0, 3: 0, 2: 0, 1: 0 };
    allVals.forEach(v => {
      const rounded = Math.round(v);
      if (counts[rounded] !== undefined) counts[rounded]++;
    });
    return [
      { name: '매우만족(5)', value: counts[5] },
      { name: '만족(4)', value: counts[4] },
      { name: '보통(3)', value: counts[3] },
      { name: '불만족(2)', value: counts[2] },
      { name: '매우불만족(1)', value: counts[1] },
    ].filter(d => d.value > 0);
  };

  const generateReport = async () => {
    if (!data) return;
    setIsGeneratingReport(true);
    try {
      const sheet1 = data.sheets[0];
      const sheet2 = data.sheets[1] || sheet1;
      
      const getSheetStats = (sheet: SheetData) => {
        return {
          effect: getAverage(sheet.rows, sheet.categories.effect),
          curriculum: getAverage(sheet.rows, sheet.categories.curriculum),
          operation: getAverage(sheet.rows, sheet.categories.operation),
          count: sheet.rows.length
        };
      };

      const stats1 = getSheetStats(sheet1);
      const stats2 = getSheetStats(sheet2);

      const prompt = `
        당신은 교육 컨설팅 전문가입니다. 다음 데이터를 바탕으로 임원 보고용 상세 분석 보고서를 작성하세요.
        전체 분량은 A4 용지 1페이지 정도의 충실한 내용(약 1,500~2,000자 내외)으로 구성해 주세요.
        언어: 한국어
        톤: 매우 전문적이고 분석적이며 정중한 비즈니스 톤
        
        중요 배경지식:
        - 각 시트(교육 과정)는 서로 독립적인 교육이며, 교육생도 다릅니다. (연속된 과정이 아님)
        - 따라서 각 과정을 개별적인 성과 지표로 분석해야 합니다.
        
        데이터 요약:
        - 교육 과정 1 (${sheet1.name}): 응답자 ${stats1.count}명, 교육효과 ${stats1.effect.toFixed(2)}, 교과편성 ${stats1.curriculum.toFixed(2)}, 교육운영 ${stats1.operation.toFixed(2)}
        - 교육 과정 2 (${sheet2.name}): 응답자 ${stats2.count}명, 교육효과 ${stats2.effect.toFixed(2)}, 교과편성 ${stats2.curriculum.toFixed(2)}, 교육운영 ${stats2.operation.toFixed(2)}
        
        보고서 구성 및 요청사항:
        [SECTION_1] 교육 과정 1 (${sheet1.name}) 성과 분석: 
        - 해당 과정의 독립적인 교육 성과 및 학습자 만족도 지표 해석
        - 항목별(효과, 편성, 운영) 강점과 개선 필요 점 도출
        
        [SECTION_2] 교육 과정 2 (${sheet2.name}) 성과 분석: 
        - 해당 과정의 독립적인 교육 성과 및 학습자 만족도 지표 해석
        - 과정 1과는 별개의 독립된 세션으로서의 특징 및 성과 분석
        
        [SECTION_3] 종합 평가 및 향후 전략적 제언: 
        - 두 교육 과정의 전반적인 운영 수준 평가
        - 데이터에 근거한 구체적인 향후 교육 설계 개선 방안(Action Item) 및 전략적 제언
        
        작성 가이드:
        - 각 섹션 시작 시 반드시 [SECTION_1], [SECTION_2], [SECTION_3] 태그를 포함해 주세요.
        - 개조식(불렛포인트)과 서술식을 적절히 혼용하여 가독성과 전문성을 동시에 확보하세요.
        - 단순 수치 나열보다는 그 수치가 의미하는 '인사이트'를 도출하는 데 집중하세요.
      `;

      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: prompt,
      });

      const fullText = response.text || "";
      
      const sections = {
        week1: fullText.split('[SECTION_1]')[1]?.split('[SECTION_2]')[0]?.trim() || "분석 중 오류가 발생했습니다.",
        week2: fullText.split('[SECTION_2]')[1]?.split('[SECTION_3]')[0]?.trim() || "분석 중 오류가 발생했습니다.",
        total: fullText.split('[SECTION_3]')[1]?.trim() || "분석 중 오류가 발생했습니다."
      };

      setExecutiveReport(sections);
      setActiveTab('report');
    } catch (error) {
      console.error("Report generation failed:", error);
    } finally {
      setIsGeneratingReport(false);
    }
  };

  if (!data) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center p-8 bg-[var(--color-bg)]">
        <div className="max-w-2xl w-full space-y-8 text-center">
          <div className="space-y-4">
            <h1 className="text-6xl font-bold tracking-tighter uppercase italic font-serif">
              Excel Insight
            </h1>
            <p className="text-xl opacity-60 font-mono">
              Upload your spreadsheet to generate a technical report.
            </p>
          </div>

          <div
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={onDrop}
            className={cn(
              "relative border-2 border-dashed border-[var(--color-line)] p-16 transition-all duration-300 group cursor-pointer",
              isDragging ? "bg-[var(--color-ink)] text-[var(--color-bg)]" : "hover:bg-[var(--color-ink)] hover:text-[var(--color-bg)]"
            )}
            onClick={() => document.getElementById('fileInput')?.click()}
          >
            <input
              id="fileInput"
              type="file"
              className="hidden"
              accept=".xlsx, .xls, .csv"
              onChange={onFileChange}
            />
            <div className="flex flex-col items-center space-y-4">
              <Upload className="w-12 h-12" />
              <p className="text-lg font-mono uppercase tracking-widest">
                {isDragging ? "Drop File Now" : "Drag & Drop or Click to Upload"}
              </p>
              <p className="text-xs opacity-50 uppercase">XLSX, XLS, CSV supported</p>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const renderReport = () => {
    if (!executiveReport) {
      return (
        <div className="flex flex-col items-center justify-center h-[60vh] space-y-6">
          <div className="p-8 border border-dashed border-[var(--color-line)] text-center space-y-4">
            <Sparkles className="w-12 h-12 mx-auto opacity-20" />
            <h3 className="text-xl font-serif italic uppercase">Generate Executive Summary</h3>
            <p className="text-sm font-mono opacity-50 max-w-md">
              AI will analyze the satisfaction data and trends to create a professional report for executives.
            </p>
            <button
              onClick={generateReport}
              disabled={isGeneratingReport}
              className="px-8 py-3 bg-[var(--color-ink)] text-[var(--color-bg)] font-mono uppercase text-xs tracking-widest hover:opacity-90 disabled:opacity-50 flex items-center space-x-2 mx-auto"
            >
              {isGeneratingReport ? (
                <>
                  <div className="w-3 h-3 border-2 border-[var(--color-bg)] border-t-transparent animate-spin rounded-full" />
                  <span>Analyzing Data...</span>
                </>
              ) : (
                <>
                  <Sparkles className="w-4 h-4" />
                  <span>Generate Report</span>
                </>
              )}
            </button>
          </div>
        </div>
      );
    }

    return (
      <div className="max-w-4xl mx-auto space-y-12 pb-20 print:p-0 print:space-y-8">
        <div className="flex justify-between items-end border-b-2 border-[var(--color-ink)] pb-4 print:border-b">
          <div>
            <h1 className="text-4xl font-serif italic uppercase tracking-tight print:text-2xl">Executive Analysis Report</h1>
            <p className="text-xs font-mono opacity-50 uppercase mt-2">Strategic Education Consulting • {new Date().toLocaleDateString()}</p>
          </div>
          <div className="flex space-x-2 print:hidden">
            <button onClick={() => window.print()} className="p-2 border border-[var(--color-line)] hover:bg-black/5 flex items-center space-x-2 text-[10px] font-mono uppercase">
              <Printer className="w-3 h-3" />
              <span>Print A4</span>
            </button>
            <button className="p-2 border border-[var(--color-line)] hover:bg-black/5"><Download className="w-4 h-4" /></button>
          </div>
        </div>

        <div className="grid grid-cols-1 gap-12 print:gap-8">
          <section className="space-y-4 break-inside-avoid">
            <div className="flex items-center space-x-4">
              <span className="text-5xl font-serif italic opacity-10 print:text-3xl">01</span>
              <h2 className="text-2xl font-bold uppercase tracking-tighter border-b border-[var(--color-line)] flex-1 pb-2 print:text-lg">
                교육 과정 1 ({data.sheets[0]?.name}) 성과 분석
              </h2>
            </div>
            <div className="text-sm text-gray-700 font-sans leading-relaxed whitespace-pre-wrap pl-12 print:pl-6 print:text-xs">
              {executiveReport.week1}
            </div>
          </section>

          <section className="space-y-4 break-inside-avoid">
            <div className="flex items-center space-x-4">
              <span className="text-5xl font-serif italic opacity-10 print:text-3xl">02</span>
              <h2 className="text-2xl font-bold uppercase tracking-tighter border-b border-[var(--color-line)] flex-1 pb-2 print:text-lg">
                교육 과정 2 ({data.sheets[1]?.name || '추가 과정'}) 성과 분석
              </h2>
            </div>
            <div className="text-sm text-gray-700 font-sans leading-relaxed whitespace-pre-wrap pl-12 print:pl-6 print:text-xs">
              {executiveReport.week2}
            </div>
          </section>

          <section className="space-y-4 break-inside-avoid">
            <div className="flex items-center space-x-4">
              <span className="text-5xl font-serif italic opacity-10 print:text-3xl">03</span>
              <h2 className="text-2xl font-bold uppercase tracking-tighter border-b border-[var(--color-line)] flex-1 pb-2 print:text-lg">통합 성과 평가 및 전략적 제언</h2>
            </div>
            <div className="bg-[var(--color-ink)] text-[var(--color-bg)] p-8 ml-12 print:ml-6 print:p-4 print:bg-gray-100 print:text-black print:border">
              <div className="text-sm text-gray-200 font-sans leading-relaxed whitespace-pre-wrap print:text-black print:text-xs">
                {executiveReport.total}
              </div>
            </div>
          </section>
        </div>
        
        <div className="hidden print:block text-center pt-12 text-[10px] font-mono opacity-30">
          CONFIDENTIAL • INTERNAL USE ONLY • PAGE 1 OF 1
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-[var(--color-bg)] flex flex-col">
      {/* Header */}
      <header className="border-b border-[var(--color-line)] p-6 flex justify-between items-center">
        <div className="flex items-center space-x-4">
          <h2 className="text-2xl font-bold italic font-serif uppercase tracking-tight">
            Report: {currentSheet?.name}
          </h2>
          {data.sheets.length > 1 && (
            <div className="flex items-center space-x-1 bg-black/5 p-1 rounded">
              {data.sheets.map((sheet, idx) => (
                <button
                  key={idx}
                  onClick={() => {
                    setData({ ...data, activeSheetIndex: idx });
                    setSelectedWeek(null);
                  }}
                  className={cn(
                    "px-3 py-1 text-[10px] font-mono uppercase transition-all",
                    data.activeSheetIndex === idx ? "bg-[var(--color-ink)] text-[var(--color-bg)]" : "hover:bg-black/10"
                  )}
                >
                  {sheet.name}
                </button>
              ))}
            </div>
          )}
        </div>
        <div className="flex items-center space-x-4">
          <div className="text-[10px] font-mono uppercase opacity-50">
            {currentSheet?.rows.length} Records
          </div>
          <button 
            onClick={() => setData(null)}
            className="p-2 hover:bg-[var(--color-ink)] hover:text-[var(--color-bg)] transition-colors border border-[var(--color-line)]"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      </header>

      <div className="flex flex-1 overflow-hidden">
        {/* Sidebar */}
        <aside className="w-64 border-r border-[var(--color-line)] flex flex-col">
          <nav className="flex-1 p-4 space-y-2">
            <button
              onClick={() => setActiveTab('overview')}
              className={cn(
                "w-full flex items-center space-x-3 p-3 text-xs font-mono uppercase tracking-wider transition-all",
                activeTab === 'overview' ? "bg-[var(--color-ink)] text-[var(--color-bg)]" : "hover:bg-black/5"
              )}
            >
              <BarChart3 className="w-4 h-4" />
              <span>Overview</span>
            </button>
            <button
              onClick={() => setActiveTab('table')}
              className={cn(
                "w-full flex items-center space-x-3 p-3 text-xs font-mono uppercase tracking-wider transition-all",
                activeTab === 'table' ? "bg-[var(--color-ink)] text-[var(--color-bg)]" : "hover:bg-black/5"
              )}
            >
              <TableIcon className="w-4 h-4" />
              <span>Data Grid</span>
            </button>
            <button
              onClick={() => setActiveTab('report')}
              className={cn(
                "w-full flex items-center space-x-3 p-3 text-xs font-mono uppercase tracking-wider transition-all",
                activeTab === 'report' ? "bg-[var(--color-ink)] text-[var(--color-bg)]" : "hover:bg-black/5"
              )}
            >
              <FileText className="w-4 h-4" />
              <span>Executive Report</span>
            </button>
            <div className="pt-4 mt-4 border-t border-[var(--color-line)]/20">
              <button
                onClick={() => setShowDiagnostics(!showDiagnostics)}
                className={cn(
                  "w-full flex items-center justify-between p-3 text-[10px] font-mono uppercase tracking-wider transition-all border border-[var(--color-line)]",
                  showDiagnostics ? "bg-yellow-100" : "hover:bg-black/5"
                )}
              >
                <span>Diagnostics</span>
                <ChevronRight className={cn("w-3 h-3 transition-transform", showDiagnostics && "rotate-90")} />
              </button>
            </div>
          </nav>
          
          <div className="p-6 border-t border-[var(--color-line)] space-y-4">
            <div className="space-y-1">
              <p className="text-[10px] opacity-50 uppercase font-mono">Total Records</p>
              <p className="text-2xl font-mono">{currentSheet?.summary.totalRows}</p>
            </div>
            <div className="space-y-1">
              <p className="text-[10px] opacity-50 uppercase font-mono">Data Points</p>
              <p className="text-2xl font-mono">{(currentSheet?.summary.totalRows || 0) * (currentSheet?.summary.totalCols || 0)}</p>
            </div>
          </div>
        </aside>

        {/* Main Content */}
        <main className="flex-1 overflow-y-auto p-8 bg-white/30">
          {activeTab === 'report' && renderReport()}
          
          {showDiagnostics && (
            <div className="mb-8 p-6 border-2 border-yellow-400 bg-yellow-50/50 space-y-4">
              <div className="flex justify-between items-center">
                <h3 className="text-sm font-bold uppercase font-mono">Data Diagnostics</h3>
                <button onClick={() => setShowDiagnostics(false)}><X className="w-4 h-4" /></button>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-[10px] font-mono">
                <div className="space-y-2">
                  <p className="font-bold border-b border-yellow-400 pb-1">Detected Headers ({currentSheet?.headers.length})</p>
                  <ul className="list-disc list-inside opacity-70">
                    {currentSheet?.headers.map((h, i) => (
                      <li key={i}>{h}</li>
                    ))}
                  </ul>
                </div>
                <div className="space-y-2">
                  <p className="font-bold border-b border-yellow-400 pb-1">Column Analysis</p>
                  <div className="space-y-1">
                    <p>Numeric Columns: {numericCols.join(', ') || 'None'}</p>
                    <p>String Columns: {stringCols.join(', ') || 'None'}</p>
                    <div className="mt-4 p-2 bg-white/50 border border-yellow-300">
                      <p className="font-bold mb-1">Sample Data (First Row):</p>
                      {currentSheet && currentSheet.rows.length > 0 ? (
                        <div className="space-y-1 overflow-x-auto">
                          {Object.entries(currentSheet.rows[0]).slice(0, 5).map(([k, v]) => (
                            <div key={k} className="flex justify-between border-b border-yellow-200 py-1">
                              <span className="opacity-60">{k}:</span>
                              <span className="font-bold">{String(v)} ({typeof v})</span>
                            </div>
                          ))}
                        </div>
                      ) : <p>No data rows found.</p>}
                    </div>
                    <p className="mt-2 text-red-600">Tip: Merged cells in Excel can result in "null" values in subsequent rows. Ensure headers are in a single row.</p>
                  </div>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'overview' && (
            <div className="space-y-8">
              {/* Week Filter */}
              {uniqueWeeks.length > 1 && (
                <div className="flex flex-wrap gap-2 p-4 border border-[var(--color-line)] bg-[var(--color-bg)]">
                  <p className="w-full text-[10px] font-mono uppercase opacity-50 mb-2">Select Week</p>
                  {uniqueWeeks.map(week => (
                    <button
                      key={week}
                      onClick={() => setSelectedWeek(week)}
                      className={cn(
                        "px-4 py-2 text-[10px] font-mono uppercase border border-[var(--color-line)] transition-all",
                        selectedWeek === week ? "bg-[var(--color-ink)] text-[var(--color-bg)]" : "hover:bg-black/5"
                      )}
                    >
                      {week}
                    </button>
                  ))}
                  <button
                    onClick={() => setSelectedWeek(null)}
                    className={cn(
                      "px-4 py-2 text-[10px] font-mono uppercase border border-[var(--color-line)] transition-all",
                      selectedWeek === null ? "bg-[var(--color-ink)] text-[var(--color-bg)]" : "hover:bg-black/5"
                    )}
                  >
                    All Weeks
                  </button>
                </div>
              )}

              {/* Top Summary Cards */}
              <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
                  {[
                    { label: '교육효과', key: currentSheet?.categories.effect, color: '#3B82F6' },
                    { label: '교과편성', key: currentSheet?.categories.curriculum, color: '#10B981' },
                    { label: '교육운영', key: currentSheet?.categories.operation, color: '#F59E0B' },
                    { label: '전체 평균', isTotal: true, color: '#6366F1' }
                  ].map((item, idx) => {
                    let val = 0;
                    if (item.isTotal) {
                      const keys = [currentSheet?.categories.effect, currentSheet?.categories.curriculum, currentSheet?.categories.operation].filter(Boolean) as string[];
                      const allVals = filteredRows.flatMap(r => keys.map(k => r[k])).filter(v => typeof v === 'number');
                      val = allVals.length ? allVals.reduce((a, b) => a + b, 0) / allVals.length : 0;
                    } else {
                      val = getAverage(filteredRows, item.key || null);
                    }
                  
                  return (
                    <div key={idx} className="border border-[var(--color-line)] p-6 bg-[var(--color-bg)]">
                      <p className="text-[10px] uppercase font-mono opacity-50 mb-2">{item.label}</p>
                      <p className="text-4xl font-mono tracking-tighter" style={{ color: item.color }}>{val.toFixed(2)}</p>
                      <div className="w-full bg-black/10 h-1 mt-4 overflow-hidden">
                        <div className="h-full" style={{ width: `${(val / 5) * 100}%`, backgroundColor: item.color }} />
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Comparison & Radar & Distribution */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                {/* Comparison W1 vs W2 */}
                {uniqueWeeks.length >= 2 && (
                  <div className="border border-[var(--color-line)] p-6 bg-[var(--color-bg)]">
                    <h3 className="text-sm font-bold italic font-serif uppercase tracking-tight mb-6">주차별 비교 (W1 → W2)</h3>
                    <div className="space-y-4">
                      {[
                        { label: '교육효과', key: currentSheet?.categories.effect },
                        { label: '교과편성', key: currentSheet?.categories.curriculum },
                        { label: '교육운영', key: currentSheet?.categories.operation }
                      ].map((item, idx) => {
                        if (!item.key) return null;
                        const w1Rows = currentSheet?.rows.filter(r => String(r[weekKey] || '') === uniqueWeeks[0]) || [];
                        const w2Rows = currentSheet?.rows.filter(r => String(r[weekKey] || '') === uniqueWeeks[1]) || [];
                        const v1 = getAverage(w1Rows, item.key);
                        const v2 = getAverage(w2Rows, item.key);
                        const diff = v2 - v1;
                        
                        return (
                          <div key={idx} className="flex items-center justify-between border-b border-[var(--color-line)]/10 pb-3 last:border-0">
                            <div className="space-y-0.5">
                              <p className="text-[9px] font-mono uppercase opacity-40">{item.label}</p>
                              <div className="flex items-baseline space-x-2">
                                <span className="text-lg font-mono">{v1.toFixed(2)}</span>
                                <span className="text-[10px] opacity-20 font-mono">→</span>
                                <span className="text-lg font-mono">{v2.toFixed(2)}</span>
                              </div>
                            </div>
                            <div className={cn(
                              "flex items-center space-x-1 px-2 py-0.5 font-mono text-[10px] uppercase",
                              diff > 0 ? "text-green-600 bg-green-50" : diff < 0 ? "text-red-600 bg-red-50" : "text-gray-500 bg-gray-50"
                            )}>
                              {diff > 0 ? <TrendingUp className="w-3 h-3" /> : diff < 0 ? <TrendingDown className="w-3 h-3" /> : <Minus className="w-3 h-3" />}
                              <span>{diff > 0 ? '+' : ''}{diff.toFixed(2)}</span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {/* Radar Chart - Explicitly uses filteredRows */}
                <div className="border border-[var(--color-line)] p-6 bg-[var(--color-bg)] flex flex-col items-center col-span-1 lg:col-span-2">
                  <h3 className="text-sm font-bold italic font-serif uppercase tracking-tight mb-4 self-start">
                    만족도 밸런스 {selectedWeek ? `(${selectedWeek})` : '(전체 평균)'}
                  </h3>
                  <div className="h-[300px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <RadarChart cx="50%" cy="50%" outerRadius="80%" data={[
                        { subject: '교육효과', A: getAverage(filteredRows, currentSheet?.categories.effect || null), fullMark: 5 },
                        { subject: '교과편성', A: getAverage(filteredRows, currentSheet?.categories.curriculum || null), fullMark: 5 },
                        { subject: '교육운영', A: getAverage(filteredRows, currentSheet?.categories.operation || null), fullMark: 5 },
                      ]}>
                        <PolarGrid stroke="rgba(0,0,0,0.1)" />
                        <PolarAngleAxis dataKey="subject" tick={{ fontSize: 10, fontFamily: 'monospace', textTransform: 'uppercase' }} />
                        <PolarRadiusAxis angle={30} domain={[0, 5]} tick={{ fontSize: 8, fontFamily: 'monospace' }} />
                        <Radar 
                          name="Satisfaction" 
                          dataKey="A" 
                          stroke="#6366F1" 
                          fill="#6366F1" 
                          fillOpacity={0.6} 
                          label={{ 
                            position: 'top', 
                            fill: '#6366F1', 
                            fontSize: 11, 
                            fontFamily: 'monospace',
                            formatter: (v: number) => v.toFixed(2)
                          }}
                        />
                        <Tooltip 
                          formatter={(value: number) => value.toFixed(2)}
                          contentStyle={{ backgroundColor: 'var(--color-bg)', border: '1px solid var(--color-line)', fontFamily: 'monospace', fontSize: '10px' }} 
                        />
                      </RadarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>

              {/* Weekly Trends - Separate Line Charts */}
              <div className="space-y-6">
                <div className="flex justify-between items-end">
                  <div>
                    <h3 className="text-lg font-bold italic font-serif uppercase tracking-tight">항목별 만족도 추이</h3>
                    <p className="text-[10px] font-mono opacity-50 uppercase">Separate Weekly Trends by Category</p>
                  </div>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  {[
                    { label: '교육효과', key: currentSheet?.categories.effect, color: '#3B82F6' },
                    { label: '교과편성', key: currentSheet?.categories.curriculum, color: '#10B981' },
                    { label: '교육운영', key: currentSheet?.categories.operation, color: '#F59E0B' }
                  ].map((cat, idx) => (
                    <div key={idx} className="border border-[var(--color-line)] p-6 bg-[var(--color-bg)]">
                      <p className="text-[10px] font-mono uppercase opacity-50 mb-4">{cat.label}</p>
                      <div className="h-[200px] w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <LineChart data={(() => {
                            const groups: any = {};
                            currentSheet?.rows.forEach(row => {
                              const w = String(row[weekKey] || '').trim();
                              // Skip if the week name is in blacklist or empty
                              if (!w || WEEK_BLACKLIST.includes(w)) return;
                              
                              if (!groups[w]) groups[w] = { week: w, vals: [] };
                              if (cat.key && typeof row[cat.key] === 'number') groups[w].vals.push(row[cat.key]);
                            });
                            
                            // Sort weeks
                            const sortedWeeks = Object.values(groups)
                              .map((g: any) => ({
                                week: g.week,
                                value: g.vals.length ? g.vals.reduce((a: any, b: any) => a + b, 0) / g.vals.length : 0
                              }))
                              .sort((a, b) => String(a.week).localeCompare(String(b.week), undefined, { numeric: true }));
                            
                            return sortedWeeks;
                          })()}>
                            <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" vertical={false} />
                            <XAxis 
                              dataKey="week" 
                              tick={{ fontSize: 8, fontFamily: 'monospace' }} 
                              stroke="var(--color-ink)"
                            />
                            <YAxis domain={[0, 5]} tick={{ fontSize: 8, fontFamily: 'monospace' }} stroke="var(--color-ink)" />
                            <Tooltip 
                              formatter={(value: number) => value.toFixed(2)}
                              contentStyle={{ backgroundColor: 'var(--color-bg)', border: '1px solid var(--color-line)', fontFamily: 'monospace', fontSize: '10px' }} 
                            />
                            <Line 
                              type="monotone" 
                              dataKey="value" 
                              stroke={cat.color} 
                              strokeWidth={2} 
                              dot={{ r: 3, fill: cat.color }} 
                              activeDot={{ r: 5 }}
                              label={{ position: 'top', fill: cat.color, fontSize: 9, fontFamily: 'monospace', formatter: (v: number) => v.toFixed(2) }}
                            />
                          </LineChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Course Comparison - Split into separate charts */}
              {currentSheet?.categories.course && (
                <div className="space-y-6">
                  <div className="border-b border-[var(--color-line)] pb-4">
                    <h3 className="text-lg font-bold italic font-serif uppercase tracking-tight">과정별 만족도 비교 {selectedWeek ? `(${selectedWeek})` : '(전체 평균)'}</h3>
                    <p className="text-[10px] font-mono opacity-50 uppercase">Satisfaction Comparison by Course Program</p>
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                    {[
                      { label: '교육효과', key: currentSheet.categories.effect, color: '#3B82F6' },
                      { label: '교과편성', key: currentSheet.categories.curriculum, color: '#10B981' },
                      { label: '교육운영', key: currentSheet.categories.operation, color: '#F59E0B' }
                    ].map((cat, idx) => (
                      <div key={idx} className="border border-[var(--color-line)] p-6 bg-[var(--color-bg)]">
                        <p className="text-[10px] font-mono uppercase opacity-50 mb-4">{cat.label} 비교</p>
                        <div className="h-[250px] w-full">
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart 
                              data={(() => {
                                const courseKey = currentSheet.categories.course!;
                                const groups: any = {};
                                const rowsToUse = selectedWeek ? filteredRows : currentSheet.rows;
                                
                                rowsToUse.forEach(row => {
                                  const c = row[courseKey] || 'N/A';
                                  if (!groups[c]) groups[c] = { course: c, vals: [] };
                                  if (cat.key && typeof row[cat.key] === 'number') groups[c].vals.push(row[cat.key]);
                                });
                                
                                return Object.values(groups).map((g: any) => ({
                                  course: g.course,
                                  value: g.vals.length ? g.vals.reduce((a: any, b: any) => a + b, 0) / g.vals.length : 0
                                })).sort((a, b) => b.value - a.value);
                              })()}
                              layout="vertical"
                              margin={{ left: 20, right: 30 }}
                            >
                              <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" horizontal={false} />
                              <XAxis type="number" domain={[0, 5]} hide />
                              <YAxis 
                                dataKey="course" 
                                type="category" 
                                tick={{ fontSize: 8, fontFamily: 'monospace' }} 
                                width={80}
                                stroke="var(--color-ink)"
                              />
                              <Tooltip 
                                formatter={(value: number) => value.toFixed(2)}
                                contentStyle={{ backgroundColor: 'var(--color-bg)', border: '1px solid var(--color-line)', fontFamily: 'monospace', fontSize: '10px' }} 
                              />
                              <Bar 
                                dataKey="value" 
                                fill={cat.color} 
                                label={{ position: 'right', fill: '#141414', fontSize: 9, fontFamily: 'monospace', formatter: (v: number) => v.toFixed(2) }} 
                              />
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}

          {activeTab === 'table' && (
            <div className="border border-[var(--color-line)] bg-[var(--color-bg)] overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="border-b border-[var(--color-line)]">
                      {currentSheet?.headers.map(header => (
                        <th key={header} className="p-4 text-[10px] font-serif italic uppercase tracking-widest opacity-50 border-r border-[var(--color-line)] last:border-r-0">
                          {header}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {currentSheet?.rows.slice(0, 50).map((row, i) => (
                      <tr key={i} className="border-b border-[var(--color-line)] last:border-b-0 hover:bg-[var(--color-ink)] hover:text-[var(--color-bg)] transition-colors group">
                        {currentSheet?.headers.map(header => (
                          <td key={header} className="p-4 text-xs font-mono border-r border-[var(--color-line)] last:border-r-0">
                            {String(row[header] ?? '')}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {(currentSheet?.rows.length || 0) > 50 && (
                <div className="p-4 text-center border-t border-[var(--color-line)]">
                  <p className="text-[10px] font-mono uppercase opacity-50">Showing first 50 of {currentSheet?.rows.length} records</p>
                </div>
              )}
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
