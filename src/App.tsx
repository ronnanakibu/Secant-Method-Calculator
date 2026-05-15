import React, { useState, useEffect, useMemo } from 'react';
import * as math from 'mathjs';
import * as XLSX from 'xlsx';
import { motion, AnimatePresence } from 'motion/react';
import { 
  LineChart, 
  Line, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip as RechartsTooltip, 
  ResponsiveContainer,
  ReferenceLine
} from 'recharts';
import { 
  Calculator, 
  Settings2, 
  Table as TableIcon, 
  LineChart as ChartIcon, 
  AlertCircle,
  Play,
  RotateCcw,
  ChevronUp,
  ChevronDown,
  Info,
  Download,
  Check
} from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// Utility for Tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// Custom e constant as requested by user
const CUSTOM_E = 2.718282;

// Formatter to avoid scientific notation (1E-x)
const formatFullNumber = (num: number, decimals: number = 10) => {
  if (isNaN(num)) return "0";
  // We use fixed to avoid E- notation, but trim trailing zeros
  const fixed = num.toFixed(decimals);
  return fixed.replace(/\.?0+$/, "");
};

// Excel Formula Converter
const convertToExcel = (expr: string, cellRef: string = "A2") => {
  try {
    let excel = expr
      .replace(/–|—/g, '-') // Normalize dashes
      .replace(/=/g, '')     // Remove equals
      .replace(/0/g, (match, offset) => {
        // Remove trailing "= 0" if input as "expr = 0"
        return offset > expr.length - 3 ? '' : '0';
      })
      .trim();

    // Handle e^something -> EXP(something)
    // Common case: e^x or e^(...)
    excel = excel.replace(/e\^(\((?:[^)(]+|\((?:[^)(]+|\([^)(]*\))*\))*\))/g, 'EXP($1)');
    excel = excel.replace(/e\^([a-zA-Z0-9.]+)/g, 'EXP($1)');
    
    // Replace any remaining 'e' (like just 'e' or 'e * x')
    excel = excel.replace(/\be\b/g, CUSTOM_E.toString());

    // Insert implicit multiplications for Excel (e.g., 5x -> 5*A1)
    excel = excel.replace(/(\d)([a-zA-Z(])/g, '$1*$2');
    excel = excel.replace(/([a-zA-Z)])(\d)/g, '$1*$2');

    // Replace x with the cell reference
    excel = excel.replace(/\bx\b/g, cellRef);

    return `=${excel}`;
  } catch (err) {
    return "Gagal mengonversi ke formula Excel.";
  }
};

interface IterationResult {
  iteration: number;
  x: number;
  fx: number;
  error: number | null;
  // Metadata for detail view formulas
  prevX?: number;
  prevPrevX?: number;
  prevFx?: number;
  prevPrevFx?: number;
  isInitial?: boolean;
}

export default function App() {
  // Form State
  const [equation, setEquation] = useState(() => localStorage.getItem('secant_eqn') || 'x^2 - e^x - 3');
  const [x0, setX0] = useState(() => localStorage.getItem('secant_x0') || '-2');
  const [x1, setX1] = useState(() => localStorage.getItem('secant_x1') || '-1');
  const [tolerance, setTolerance] = useState(() => localStorage.getItem('secant_tol') || '0.0001');
  const [maxIterations, setMaxIterations] = useState(() => localStorage.getItem('secant_max') || '20');
  const [precision, setPrecision] = useState(() => localStorage.getItem('secant_prec') || '6');
  
  // Persist state to localStorage
  useEffect(() => {
    localStorage.setItem('secant_eqn', equation);
    localStorage.setItem('secant_x0', x0);
    localStorage.setItem('secant_x1', x1);
    localStorage.setItem('secant_tol', tolerance);
    localStorage.setItem('secant_max', maxIterations);
    localStorage.setItem('secant_prec', precision);
  }, [equation, x0, x1, tolerance, maxIterations, precision]);

  // Calculation State
  const [results, setResults] = useState<IterationResult[]>([]);
  const [errorStatus, setErrorStatus] = useState<string | null>(null);
  const [finalRoot, setFinalRoot] = useState<number | null>(null);
  const [selectedIter, setSelectedIter] = useState<IterationResult | null>(null);

  // Excel Builder State
  const [builderFormula, setBuilderFormula] = useState("");

  // Helper to add term to builder
  const addToBuilder = (term: string) => {
    setBuilderFormula(prev => prev + term);
  };
  
// Custom Math evaluator with the user's specific e
const evaluateFunction = (expr: string, x: number, precision: string) => {
  try {
    let cleanExpr = expr
      .replace(/–|—/g, '-') 
      .replace(/=/g, '')     
      .replace(/=\s*0\s*$/, '')
      .trim();

    // Fix precedence for e^ terms: e^2x -> e^(2x)
    cleanExpr = cleanExpr.replace(/e\^([-+.\w]+)/g, (match, p1) => {
      if (p1.startsWith('(') && p1.endsWith(')')) return match;
      return `e^(${p1})`;
    });

    const scope = { 
      x, 
      e: CUSTOM_E,
      exp: (val: number) => Math.pow(CUSTOM_E, val) 
    };

    cleanExpr = cleanExpr.replace(/(\d)([a-zA-Z(])/g, '$1*$2');
    
    const node = math.parse(cleanExpr);
    const code = node.compile();
    let result = code.evaluate(scope);
    
    if (typeof result !== 'number' || !isFinite(result)) {
      throw new Error("Hasil bukan angka valid.");
    }

    const p = parseInt(precision);
    if (!isNaN(p)) {
      result = Number(result.toFixed(p));
    }

    return result;
  } catch (err: any) {
    throw new Error(`Persamaan tidak valid: ${err.message || 'Cek penulisan formula.'}`);
  }
};

const DetailEvaluation = ({ selectedIter, equation, precision, CUSTOM_E }: { selectedIter: IterationResult, equation: string, precision: string, CUSTOM_E: number }) => (
  <div className="border-t border-indigo-500/20 pt-6 mt-4">
    <div className="flex items-center gap-2 mb-3">
      <div className="w-1.5 h-1.5 bg-indigo-400 rounded-full"></div>
      <span className="text-indigo-300 font-bold text-[10px] uppercase tracking-widest">Detail Evaluasi Fungsi f(x_{selectedIter.iteration})</span>
    </div>
    
    <div className="space-y-4 bg-slate-900/50 p-4 rounded-xl border border-indigo-500/10">
      <div className="flex flex-col gap-3">
        <div className="flex items-start gap-3">
          <div className="w-5 h-5 rounded bg-indigo-500/20 flex items-center justify-center text-[10px] text-indigo-300 shrink-0 mt-0.5">A</div>
          <div>
            <p className="text-slate-500 text-[9px] uppercase tracking-wider mb-1">Persamaan Asli</p>
            <code className="text-indigo-200">f(x) = {equation}</code>
          </div>
        </div>

        <div className="flex items-start gap-3">
          <div className="w-5 h-5 rounded bg-indigo-500/20 flex items-center justify-center text-[10px] text-indigo-300 shrink-0 mt-0.5">B</div>
          <div>
            <p className="text-slate-500 text-[9px] uppercase tracking-wider mb-1">Substitusi x = {formatFullNumber(selectedIter.x, parseInt(precision))}</p>
            <code className="text-indigo-200 break-all leading-relaxed whitespace-pre-wrap">
              f({formatFullNumber(selectedIter.x, parseInt(precision))}) = {equation.replace(/(?<![a-zA-Z])x/g, `(${formatFullNumber(selectedIter.x, parseInt(precision))})`).replace(/\be\b/g, `(${CUSTOM_E})`)}
            </code>
          </div>
        </div>

        <div className="flex items-start gap-3">
          <div className="w-5 h-5 rounded bg-indigo-500/20 flex items-center justify-center text-[10px] text-indigo-300 shrink-0 mt-0.5">C</div>
          <div className="flex-1">
            <p className="text-slate-500 text-[9px] uppercase tracking-wider mb-2">Evaluasi Suku-Suku & Fungsi:</p>
            <div className="text-indigo-100 text-[10px] space-y-4">
              <div className="grid grid-cols-1 gap-2">
                {Array.from(equation.matchAll(/(?:^|(?<=[^a-zA-Z0-9^]))([+-]?\d*(?:\.\d+)?)?\*?\s*e\^([-+()\d.*]*x)/g)).map((match, idx) => {
                  const rawCoeff = match[1] || "";
                  const coeff = rawCoeff === "" || rawCoeff === "+" ? "1" : (rawCoeff === "-" ? "-1" : rawCoeff);
                  const exponent = match[2];
                  const substitutedExponent = exponent.replace(/(?<![a-zA-Z])x/g, `(${formatFullNumber(selectedIter.x, parseInt(precision))})`);
                  let val = 0;
                  try {
                    val = evaluateFunction(`${coeff} * e^(${exponent})`, selectedIter.x, precision);
                  } catch (err) {
                    val = parseFloat(coeff) * Math.exp(selectedIter.x);
                  }
                  return (
                    <div key={`exp-${idx}`} className="flex justify-between items-center border-b border-white/5 pb-1 italic">
                      <span className="text-slate-500 font-mono">
                        {coeff !== "1" && coeff !== "-1" ? `${coeff}*` : (coeff === "-1" ? "-" : "")}e^{substitutedExponent}
                      </span>
                      <span className="text-indigo-200">{val.toFixed(parseInt(precision) + 4)}</span>
                    </div>
                  );
                })}
                {Array.from(equation.matchAll(/(?:^|(?<=[^a-zA-Z0-9^]))([+-]?\d*(?:\.\d+)?)?\*?\s*x\^(\d+)/g)).map((match, idx) => {
                  const rawCoeff = match[1] || "";
                  const coeff = rawCoeff === "" || rawCoeff === "+" ? "1" : (rawCoeff === "-" ? "-1" : rawCoeff);
                  const p = parseInt(match[2]);
                  const termValue = parseFloat(coeff) * Math.pow(selectedIter.x, p);
                  return (
                    <div key={`pwr-${idx}`} className="flex justify-between items-center border-b border-white/5 pb-1 italic">
                      <span className="text-slate-500 font-mono">
                        {coeff !== "1" && coeff !== "-1" ? `${coeff}*` : (coeff === "-1" ? "-" : "")}({formatFullNumber(selectedIter.x, parseInt(precision))})^{p}
                      </span>
                      <span className="text-indigo-200">{termValue.toFixed(parseInt(precision) + 4)}</span>
                    </div>
                  );
                })}
                {Array.from(equation.matchAll(/(?:^|(?<=[^a-zA-Z0-9^]))([+-]?\d*(?:\.\d+)?)?\*?\s*x\b(?!\^)/g)).map((match, idx) => {
                  const rawCoeff = match[1] || "";
                  const coeff = rawCoeff === "" || rawCoeff === "+" ? "1" : (rawCoeff === "-" ? "-1" : rawCoeff);
                  const termValue = parseFloat(coeff) * selectedIter.x;
                  return (
                    <div key={`lin-${idx}`} className="flex justify-between items-center border-b border-white/5 pb-1 italic">
                      <span className="text-slate-500 font-mono">
                        {coeff !== "1" && coeff !== "-1" ? `${coeff}*` : (coeff === "-1" ? "-" : "")}({formatFullNumber(selectedIter.x, parseInt(precision))})
                      </span>
                      <span className="text-indigo-200">{termValue.toFixed(parseInt(precision) + 4)}</span>
                    </div>
                  );
                })}
              </div>
              <div className="mt-2 p-2 bg-teal-500/10 rounded border border-teal-500/20">
                <div className="flex justify-between items-baseline">
                  <span className="text-xs text-teal-400 font-bold uppercase tracking-tighter">f(x_{selectedIter.iteration}) =</span>
                  <span className="text-sm font-mono text-teal-300 font-bold">{formatFullNumber(selectedIter.fx, parseInt(precision))}</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
);

const DetailError = ({ selectedIter, precision, tolerance }: { selectedIter: IterationResult, precision: string, tolerance: string }) => (
  <div className="flex items-start gap-3 mt-4">
    <div className="w-5 h-5 rounded bg-pink-500/20 flex items-center justify-center text-[10px] text-pink-300 shrink-0 mt-0.5">E</div>
    <div className="flex-1">
      <p className="text-slate-500 text-[9px] uppercase tracking-wider mb-2">Penjabaran Error |x_{selectedIter.iteration} - x_{selectedIter.iteration-1}|</p>
      <div className="text-pink-100 text-[10px] space-y-4">
        <div className="space-y-1">
          <p className="text-pink-400/80 font-bold uppercase tracking-tighter text-[9px]">1. Rumus & Substitusi:</p>
          <div className="pl-3 py-1 bg-white/5 rounded border-l-2 border-pink-500/30 font-mono">
            ε = |{formatFullNumber(selectedIter.x, parseInt(precision))} - {formatFullNumber(selectedIter.prevX || 0, parseInt(precision))}|
          </div>
        </div>
        <div className="mt-2 p-2 bg-pink-500/10 rounded border border-pink-500/20">
          <div className="flex justify-between items-baseline">
            <span className="text-xs text-pink-400 font-bold">Error (ε) =</span>
            <span className="text-sm font-mono text-pink-300 font-bold">{formatFullNumber(selectedIter.error || 0, parseInt(precision))}</span>
          </div>
        </div>
        {selectedIter.error !== null && selectedIter.error < parseFloat(tolerance) && (
          <div className="text-[9px] text-teal-400 mt-1 italic flex items-center gap-1">
            <Check size={10} /> Konvergensi Tercapai (ε &lt; {tolerance})
          </div>
        )}
      </div>
    </div>
  </div>
);

const DetailNextRoot = ({ selectedIter, precision }: { selectedIter: IterationResult, precision: string }) => {
  const nextXRaw = (selectedIter.fx * ((selectedIter.x) - (selectedIter.prevX || 0))) / ((selectedIter.fx) - (selectedIter.prevFx || 0));
  const nextX = (selectedIter.x) - nextXRaw;
  
  return (
    <div className="flex items-start gap-3 mt-4">
      <div className="w-5 h-5 rounded bg-blue-500/20 flex items-center justify-center text-[10px] text-blue-300 shrink-0 mt-0.5">F</div>
      <div className="flex-1">
        <p className="text-slate-500 text-[9px] uppercase tracking-wider mb-2">Akar Estimasi Berikutnya (x_{selectedIter.iteration + 1})</p>
        <div className="text-blue-100 text-[10px] space-y-3">
          <div className="pl-3 py-1 bg-white/5 rounded border-l-2 border-blue-500/30 font-mono text-[9px] break-all leading-relaxed">
            x_{selectedIter.iteration+1} = {formatFullNumber(selectedIter.x, parseInt(precision))} - [ {formatFullNumber(selectedIter.fx, parseInt(precision))} * ({formatFullNumber(selectedIter.x, parseInt(precision))} - {formatFullNumber(selectedIter.prevX || 0, parseInt(precision))}) ] / [ {formatFullNumber(selectedIter.fx, parseInt(precision))} - {formatFullNumber(selectedIter.prevFx || 0, parseInt(precision))} ]
          </div>
          <div className="mt-1 p-2 bg-blue-500/10 rounded border border-blue-500/20">
            <div className="flex justify-between items-baseline">
              <span className="text-xs text-blue-400 font-bold">Next x =</span>
              <span className="text-sm font-mono text-blue-300 font-bold">{formatFullNumber(nextX, parseInt(precision))}</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

  const calculateSecant = () => {
    setErrorStatus(null);
    setResults([]);
    setFinalRoot(null);

    // 1. Validate Inputs
    const tol = parseFloat(tolerance);
    const maxI = parseInt(maxIterations);
    const precValue = parseInt(precision);
    // Use raw values for initial guesses
    let curX0 = parseFloat(x0);
    let curX1 = parseFloat(x1);

    if (equation.trim() === "") {
      setErrorStatus("Persamaan f(x) tidak boleh kosong.");
      return;
    }

    if (isNaN(curX0) || isNaN(curX1)) {
      setErrorStatus("Tebakan awal (x0, x1) harus berupa angka valid.");
      return;
    }

    if (isNaN(tol) || tol <= 0 || tol >= 1) {
      setErrorStatus("Toleransi harus di antara 0 dan 1 (contoh: 0.0001).");
      return;
    }

    if (isNaN(maxI) || maxI <= 0 || maxI > 1000) {
      setErrorStatus("Maksimum iterasi harus di antara 1 dan 1000.");
      return;
    }

    if (isNaN(precValue) || precValue < 0 || precValue > 15) {
      setErrorStatus("Presisi harus di antara 0 dan 15.");
      return;
    }

    if (curX0 === curX1) {
      setErrorStatus("x0 dan x1 tidak boleh sama.");
      return;
    }

    const iterativeResults: IterationResult[] = [];

    try {
      // Step 0: Initial guess x0
      const fx0 = evaluateFunction(equation, curX0, precision);
      iterativeResults.push({
        iteration: 0,
        x: curX0,
        fx: fx0,
        error: null,
        isInitial: true
      });

      // Step 1: Initial guess x1
      const fx1 = evaluateFunction(equation, curX1, precision);
      iterativeResults.push({
        iteration: 1,
        x: curX1,
        fx: fx1,
        error: Math.abs(curX1 - curX0),
        prevX: curX0,
        prevFx: fx0,
        isInitial: true
      });

      let r_m2 = curX0;
      let f_m2 = fx0;
      let r_m1 = curX1;
      let f_m1 = fx1;

      // Start calculation for r=2, 3, ...
      for (let i = 2; i <= maxI; i++) {
        const denominator = f_m1 - f_m2;
        if (Math.abs(denominator) < 1e-20) {
          setErrorStatus("Pembagi nol. Metode Secant gagal.");
          break;
        }

        const nextX = r_m1 - (f_m1 * (r_m1 - r_m2)) / denominator;
        const fxNext = evaluateFunction(equation, nextX, precision);
        const currentError = Math.abs(nextX - r_m1);

        iterativeResults.push({
          iteration: i,
          x: nextX,
          fx: fxNext,
          error: currentError,
          prevX: r_m1,
          prevPrevX: r_m2,
          prevFx: f_m1,
          prevPrevFx: f_m2,
          isInitial: false
        });

        if (currentError < tol) {
          setFinalRoot(nextX);
          break;
        }

        // Update for next iteration
        r_m2 = r_m1;
        f_m2 = f_m1;
        r_m1 = nextX;
        f_m1 = fxNext;
      }

      setResults(iterativeResults);
      if (iterativeResults.length > 0) {
        setSelectedIter(iterativeResults[iterativeResults.length - 1]);
      }
    } catch (err: any) {
      setErrorStatus(err.message);
    }
  };

  const exportToExcel = () => {
    if (results.length === 0) return;

    const data = results.map((r) => {
      return {
        "Iterasi (r)": r.iteration,
        "x r": r.x,
        "f(x r)": r.fx,
        "Error Absolut": r.error === null ? "-" : r.error,
        "Excel Formula": convertToExcel(equation, "B" + (r.iteration + 2)),
        "Status": r.isInitial ? "Initial Guess" : "Calculated Root"
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Secant Results");

    // Auto-width columns
    const maxWidths = Object.keys(data[0]).map(key => Math.max(key.length, 15));
    worksheet["!cols"] = maxWidths.map(w => ({ wch: w }));

    XLSX.writeFile(workbook, `Secant_Analysis_${new Date().getTime()}.xlsx`);
  };

  // Generate chart data
  const chartData = useMemo(() => {
    if (!finalRoot && results.length === 0) {
      // Sample range around initial guesses
      const s0 = parseFloat(x0);
      const s1 = parseFloat(x1);
      if (isNaN(s0) || isNaN(s1)) return [];
      
      const start = Math.min(s0, s1) - 2;
      const end = Math.max(s0, s1) + 2;
      const data = [];
      try {
        for (let x = start; x <= end; x += (end - start) / 50) {
          data.push({ x, y: evaluateFunction(equation, x, precision) });
        }
      } catch {}
      return data;
    }

    // Range around the results
    const allX = results.map(r => r.x);
    const start = Math.min(...allX) - 1;
    const end = Math.max(...allX) + 1;
    const data = [];
    try {
      for (let x = start; x <= end; x += (end - start) / 50) {
        data.push({ x, y: evaluateFunction(equation, x, precision) });
      }
    } catch {}
    return data;
  }, [results, finalRoot, equation, x0, x1]);

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 font-sans relative overflow-hidden flex flex-col">
      {/* Mesh Gradient Background Layers */}
      <div className="absolute top-[-10%] left-[-10%] w-[50%] h-[50%] bg-indigo-600/20 rounded-full blur-[120px] pointer-events-none"></div>
      <div className="absolute bottom-[-10%] right-[-10%] w-[50%] h-[50%] bg-teal-500/15 rounded-full blur-[120px] pointer-events-none"></div>

      {/* Header */}
      <header className="h-16 flex items-center px-8 border-b border-white/10 backdrop-blur-md z-10 sticky top-0 bg-slate-950/30">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-indigo-500 rounded-lg flex items-center justify-center font-bold text-white shadow-lg shadow-indigo-500/40">
            <Calculator size={18} />
          </div>
          <div>
            <h1 className="text-xl font-semibold tracking-tight">Secant <span className="text-indigo-400 font-light">Numerical Solver</span></h1>
          </div>
        </div>
        <div className="ml-auto flex items-center gap-6">
          <div className="px-3 py-1 bg-white/5 border border-white/10 rounded-full text-xs font-mono text-slate-400">
            e ≈ <span className="text-indigo-300">{CUSTOM_E}</span>
          </div>
          <div className="h-4 w-[1px] bg-white/10 hidden sm:block"></div>
          <span className="text-xs text-slate-500 uppercase tracking-widest hidden sm:block">Engine Optimized</span>
        </div>
      </header>

      <main className="max-w-[1400px] mx-auto p-6 grid grid-cols-1 lg:grid-cols-12 gap-6 flex-1 z-10 w-full">
        {/* Sidebar Inputs */}
        <div className="lg:col-span-4 space-y-6">
          <section className="bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 shadow-2xl">
            <div className="flex items-center gap-2 mb-6 border-b border-white/5 pb-4">
              <Settings2 size={18} className="text-indigo-400" />
              <h2 className="font-medium">Konfigurasi</h2>
            </div>
            
            <div className="space-y-4">
              <div>
                <label className="block text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-2 ml-1">
                  Persamaan f(x)
                </label>
                <div className="relative">
                  <input 
                    type="text" 
                    value={equation}
                    onChange={(e) => setEquation(e.target.value)}
                    className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-3 font-mono text-sm text-slate-100 focus:outline-none focus:ring-2 focus:ring-indigo-500/50 transition-all"
                    placeholder="Contoh: x^2 - e^x - 3"
                  />
                  <div className="absolute right-3 top-1/2 -translate-y-1/2 text-[10px] text-indigo-400/50 font-mono">f(x)</div>
                </div>
                <p className="mt-2 text-[11px] text-slate-500 italic ml-1">Gunakan 'e' untuk nilai {CUSTOM_E}</p>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-2 ml-1">
                    Tebakan x₀
                  </label>
                  <input 
                    type="text" 
                    value={x0}
                    onChange={(e) => setX0(e.target.value)}
                    className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2 text-sm font-mono text-slate-100 outline-none focus:ring-2 focus:ring-indigo-500/50"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-2 ml-1">
                    Tebakan x₁
                  </label>
                  <input 
                    type="text" 
                    value={x1}
                    onChange={(e) => setX1(e.target.value)}
                    className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2 text-sm font-mono text-slate-100 outline-none focus:ring-2 focus:ring-indigo-500/50"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-2 ml-1">
                    Toleransi (ε)
                  </label>
                  <input 
                    type="text" 
                    value={tolerance}
                    onChange={(e) => setTolerance(e.target.value)}
                    className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2 text-sm font-mono text-slate-100 outline-none focus:ring-2 focus:ring-indigo-500/50"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-2 ml-1">
                    Maks Iterasi
                  </label>
                  <input 
                    type="text" 
                    value={maxIterations}
                    onChange={(e) => setMaxIterations(e.target.value)}
                    className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2 text-sm font-mono text-slate-100 outline-none focus:ring-2 focus:ring-indigo-500/50"
                  />
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-2 ml-1">
                  Presisi Kalkulator (Desimal)
                </label>
                <input 
                  type="number" 
                  value={precision}
                  onChange={(e) => setPrecision(e.target.value)}
                  className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-3 font-mono text-sm text-slate-100 focus:outline-none focus:ring-2 focus:ring-indigo-500/50 transition-all"
                  placeholder="Contoh: 6"
                />
              </div>

              <div className="pt-4 flex gap-3">
                <button 
                  onClick={calculateSecant}
                  className="flex-1 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl py-3 font-semibold shadow-lg shadow-indigo-600/20 transition-all flex items-center justify-center gap-2 active:scale-[0.98]"
                >
                  <Play size={16} fill="currentColor" />
                  Hitung
                </button>
                <button 
                  onClick={() => {
                    setResults([]);
                    setFinalRoot(null);
                    setErrorStatus(null);
                    setSelectedIter(null);
                  }}
                  className="p-3 bg-white/5 border border-white/10 rounded-xl hover:bg-white/10 transition-colors text-slate-400"
                  title="Reset"
                >
                  <RotateCcw size={18} />
                </button>
              </div>
            </div>
          </section>

          {/* Detailed Substitution Section */}
          {results.length > 0 && selectedIter && (
            <motion.section 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-indigo-500/5 border border-indigo-500/20 backdrop-blur-xl rounded-2xl p-6 shadow-2xl overflow-hidden relative"
            >
              <div className="absolute top-0 right-0 p-4 opacity-5 pointer-events-none">
                <Info size={120} />
              </div>
              <div className="flex items-center justify-between mb-4 border-b border-indigo-500/10 pb-3">
                <div className="flex items-center gap-2">
                  <Calculator size={16} className="text-indigo-400" />
                  <h2 className="text-sm font-semibold text-indigo-100 uppercase tracking-tighter">Detail Perhitungan r = {selectedIter.iteration}</h2>
                </div>
                <div className="flex gap-1">
                  <button 
                    onClick={() => {
                      const text = `Iterasi r = ${selectedIter.iteration}\n` +
                                   `x_r = ${selectedIter.x}\n` +
                                   `f(x_r) = ${selectedIter.fx}\n` +
                                   (selectedIter.error !== null ? `Error = ${selectedIter.error}\n` : '') +
                                   `Jenis: ${selectedIter.isInitial ? 'Tebakan Awal' : 'Akar Terkalkulasi'}`;
                      navigator.clipboard.writeText(text);
                    }}
                    className="p-1 px-2 rounded bg-teal-500/10 border border-teal-500/20 text-teal-300 hover:bg-teal-500/20 transition-all text-[9px] font-bold mr-2 uppercase"
                  >
                    Copy Detail
                  </button>
                  <button 
                    onClick={() => selectedIter.iteration > 0 && setSelectedIter(results[selectedIter.iteration - 1])}
                    disabled={selectedIter.iteration === 0}
                    className="p-1 rounded bg-indigo-500/10 border border-indigo-500/20 text-indigo-300 disabled:opacity-30 hover:bg-indigo-500/20 transition-all"
                  >
                    <ChevronUp size={14} />
                  </button>
                  <button 
                    onClick={() => selectedIter.iteration < results.length - 1 && setSelectedIter(results[selectedIter.iteration + 1])}
                    disabled={selectedIter.iteration === results.length - 1}
                    className="p-1 rounded bg-indigo-500/10 border border-indigo-500/20 text-indigo-300 disabled:opacity-30 hover:bg-indigo-500/20 transition-all"
                  >
                    <ChevronDown size={14} />
                  </button>
                </div>
              </div>
              
              {selectedIter.iteration === 0 ? (
                <div className="space-y-6">
                  {/* Step 0: Evaluation only */}
                  <div className="bg-slate-900/40 p-4 rounded-xl border border-white/5 font-mono text-[11px]">
                    <p className="text-indigo-400 font-bold uppercase tracking-widest text-[9px] mb-2">Tahap Inisialisasi: Tebakan Pertama</p>
                    <p className="text-indigo-200">x₀ = {formatFullNumber(selectedIter.x, parseInt(precision))}</p>
                    <p className="text-slate-500 mt-1 italic">Mengevaluasi fungsi pada titik awal x₀...</p>
                  </div>
                  <DetailEvaluation 
                    selectedIter={selectedIter} 
                    equation={equation} 
                    precision={precision} 
                    CUSTOM_E={CUSTOM_E} 
                  />
                </div>
              ) : selectedIter.iteration === 1 ? (
                <div className="space-y-6">
                  {/* Step 1: Evaluation + Error */}
                  <div className="bg-slate-900/40 p-4 rounded-xl border border-white/5 font-mono text-[11px]">
                    <p className="text-indigo-400 font-bold uppercase tracking-widest text-[9px] mb-2">Tahap Inisialisasi: Tebakan Kedua</p>
                    <p className="text-indigo-200">x₁ = {formatFullNumber(selectedIter.x, parseInt(precision))}</p>
                    <p className="text-slate-500 mt-1 italic">Mengevaluasi fungsi pada titik awal x₁...</p>
                  </div>
                  <DetailEvaluation 
                    selectedIter={selectedIter} 
                    equation={equation} 
                    precision={precision} 
                    CUSTOM_E={CUSTOM_E} 
                  />
                  <DetailError 
                    selectedIter={selectedIter} 
                    precision={precision} 
                    tolerance={tolerance} 
                  />
                  <DetailNextRoot 
                    selectedIter={selectedIter} 
                    precision={precision} 
                  />
                </div>
              ) : (
                <div className="space-y-4 font-mono text-[11px]">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-2">
                    <div className="bg-slate-900/40 p-3 rounded-lg border border-white/5">
                      <div className="text-[9px] text-slate-500 uppercase mb-1">Input Sebelumnya</div>
                      <p>x_{selectedIter.iteration-2} = {formatFullNumber(selectedIter.prevPrevX || 0, parseInt(precision))}</p>
                      <p>x_{selectedIter.iteration-1} = {formatFullNumber(selectedIter.prevX || 0, parseInt(precision))}</p>
                    </div>
                    <div className="bg-slate-900/40 p-3 rounded-lg border border-white/5">
                      <div className="text-[9px] text-slate-500 uppercase mb-1">Fungsi Sebelumnya</div>
                      <p>f(x_{selectedIter.iteration-2}) = {formatFullNumber(selectedIter.prevPrevFx || 0, parseInt(precision))}</p>
                      <p>f(x_{selectedIter.iteration-1}) = {formatFullNumber(selectedIter.prevFx || 0, parseInt(precision))}</p>
                    </div>
                  </div>
                  <div className="space-y-3 bg-slate-900/30 p-4 rounded-xl border border-white/5 leading-relaxed overflow-x-auto">
                    <div className="border-l-2 border-indigo-500/30 pl-4 py-1">
                      <span className="text-slate-500 text-[10px] uppercase tracking-wider block mb-1">1. Rumus Secant untuk x_{selectedIter.iteration}</span>
                      <div className="text-indigo-300 font-bold">x_{selectedIter.iteration} = x_{selectedIter.iteration-1} - [f(x_{selectedIter.iteration-1}) · (x_{selectedIter.iteration-1} - x_{selectedIter.iteration-2})] / [f(x_{selectedIter.iteration-1}) - f(x_{selectedIter.iteration-2})]</div>
                    </div>
                    
                    <div className="border-l-2 border-slate-700 pl-4 py-1">
                      <span className="text-slate-500 text-[10px] uppercase tracking-wider block mb-1">2. Substitusi Nilai</span>
                      <div className="text-indigo-100 text-[10px] whitespace-pre-wrap break-all">
                        x_{selectedIter.iteration} = {formatFullNumber(selectedIter.prevX || 0, parseInt(precision))} - [{formatFullNumber(selectedIter.prevFx || 0, parseInt(precision))} · ({formatFullNumber(selectedIter.prevX || 0, parseInt(precision))} - {formatFullNumber(selectedIter.prevPrevX || 0, parseInt(precision))})] / [{formatFullNumber(selectedIter.prevFx || 0, parseInt(precision))} - {formatFullNumber(selectedIter.prevPrevFx || 0, parseInt(precision))}]
                      </div>
                    </div>

                    <div className="grid grid-cols-1 gap-3 border-t border-white/5 pt-3">
                      <div className="bg-slate-900/40 p-2 rounded border border-white/5">
                        <span className="text-[9px] text-slate-500 uppercase block mb-1">3. Hitung Pembilang (Numerator)</span>
                        <p className="text-indigo-200">
                          {formatFullNumber(selectedIter.prevFx || 0, parseInt(precision))} · {formatFullNumber((selectedIter.prevX || 0) - (selectedIter.prevPrevX || 0), parseInt(precision))} = <span className="text-indigo-400 font-bold">{formatFullNumber((selectedIter.prevFx || 0) * ((selectedIter.prevX || 0) - (selectedIter.prevPrevX || 0)), parseInt(precision))}</span>
                        </p>
                      </div>
                      
                      <div className="bg-slate-900/40 p-2 rounded border border-white/5">
                        <span className="text-[9px] text-slate-500 uppercase block mb-1">4. Hitung Penyebut (Denominator)</span>
                        <p className="text-indigo-200">
                          {formatFullNumber(selectedIter.prevFx || 0, parseInt(precision))} - {formatFullNumber(selectedIter.prevPrevFx || 0, parseInt(precision))} = <span className="text-indigo-400 font-bold">{formatFullNumber((selectedIter.prevFx || 0) - (selectedIter.prevPrevFx || 0), parseInt(precision))}</span>
                        </p>
                      </div>

                      <div className="bg-indigo-500/10 p-2 rounded border border-indigo-500/20">
                        <span className="text-[9px] text-indigo-400 uppercase block mb-1">5. Faktor Koreksi (Correction Term)</span>
                        <p className="text-teal-300">
                          {formatFullNumber((selectedIter.prevFx || 0) * ((selectedIter.prevX || 0) - (selectedIter.prevPrevX || 0)), parseInt(precision))} / {formatFullNumber((selectedIter.prevFx || 0) - (selectedIter.prevPrevFx || 0), parseInt(precision))} = <span className="font-bold underline">{formatFullNumber(((selectedIter.prevFx || 0) * ((selectedIter.prevX || 0) - (selectedIter.prevPrevX || 0))) / ((selectedIter.prevFx || 0) - (selectedIter.prevPrevFx || 0)), parseInt(precision))}</span>
                        </p>
                      </div>
                    </div>

                    <div className="border-t-2 border-teal-500/40 pt-3 flex flex-col gap-2">
                      <span className="text-teal-400 font-bold text-[10px] uppercase tracking-widest">6. Hasil Akar x_{selectedIter.iteration}</span>
                      <div className="text-teal-400 font-mono text-xs pl-2 bg-teal-500/5 p-2 rounded-lg">
                        x_{selectedIter.iteration} = {formatFullNumber(selectedIter.prevX || 0, parseInt(precision))} - {formatFullNumber(((selectedIter.prevFx || 0) * ((selectedIter.prevX || 0) - (selectedIter.prevPrevX || 0))) / ((selectedIter.prevFx || 0) - (selectedIter.prevPrevFx || 0)), parseInt(precision))}
                        <p className="mt-1 text-sm font-bold">x_{selectedIter.iteration} = {formatFullNumber(selectedIter.x, parseInt(precision))}</p>
                      </div>
                    </div>

                    <DetailEvaluation 
                      selectedIter={selectedIter} 
                      equation={equation} 
                      precision={precision} 
                      CUSTOM_E={CUSTOM_E} 
                    />
                    
                    <DetailError 
                      selectedIter={selectedIter} 
                      precision={precision} 
                      tolerance={tolerance} 
                    />

                    {selectedIter.error >= parseFloat(tolerance) && (
                      <DetailNextRoot 
                        selectedIter={selectedIter} 
                        precision={precision} 
                      />
                    )}
                  </div>
                </div>
              )}
            </motion.section>
          )}

          {/* Excel Formula Section */}
          <section className="bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 shadow-2xl">
            <div className="flex items-center gap-2 mb-4 border-b border-white/5 pb-3">
              <TableIcon size={16} className="text-teal-400" />
              <h2 className="text-sm font-medium">Excel Converter & Builder</h2>
            </div>
            
            <div className="space-y-4">
              {/* Formula Builder Buttons */}
              <div>
                <label className="block text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-2">Interactive Builder</label>
                <div className="flex flex-wrap gap-1.5 mb-3">
                  {[
                    { label: "x", value: "x" },
                    { label: "e^x", value: "e^x" },
                    { label: "+", value: " + " },
                    { label: "-", value: " - " },
                    { label: "*", value: " * " },
                    { label: "/", value: " / " },
                    { label: "^", value: "^" },
                    { label: "(", value: "(" },
                    { label: ")", value: ")" },
                    { label: "SIN", value: "sin(x)" },
                    { label: "COS", value: "cos(x)" },
                  ].map((btn) => (
                    <button
                      key={btn.label}
                      onClick={() => addToBuilder(btn.value)}
                      className="px-2 py-1 rounded bg-white/5 border border-white/10 text-[10px] font-mono hover:bg-white/10 hover:border-white/20 transition-all text-slate-300"
                    >
                      {btn.label}
                    </button>
                  ))}
                  <button
                    onClick={() => setBuilderFormula("")}
                    className="px-2 py-1 rounded bg-red-500/10 border border-red-500/20 text-[10px] font-bold hover:bg-red-500/20 transition-all text-red-400 ml-auto"
                  >
                    CLEAR
                  </button>
                  <button
                    onClick={() => setEquation(builderFormula)}
                    disabled={!builderFormula}
                    className="px-2 py-1 rounded bg-teal-500/10 border border-teal-500/20 text-[10px] font-bold hover:bg-teal-500/20 transition-all text-teal-400 disabled:opacity-30 disabled:cursor-not-allowed"
                  >
                    USE IN SECANT
                  </button>
                </div>
                
                <div className="bg-black/20 border border-white/5 rounded-lg p-2 min-h-[36px] flex items-center font-mono text-[11px] text-indigo-300">
                  {builderFormula || <span className="text-slate-600 italic">Klik tombol di atas untuk membangun persamaan...</span>}
                </div>
              </div>

              <div className="pt-2">
                <label className="block text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-2">Formula Excel (Cell A2)</label>
                <div className="bg-slate-900/80 border border-white/10 rounded-xl p-3 font-mono text-[11px] text-teal-300 break-all select-all flex justify-between items-center group">
                  <code>{convertToExcel(builderFormula || equation)}</code>
                  <div className="opacity-0 group-hover:opacity-100 transition-opacity text-[8px] text-slate-500 uppercase">Click to copy</div>
                </div>
                <p className="mt-2 text-[10px] text-slate-500 leading-relaxed">
                  Salin formula ini ke Excel. Persamaan yang sedang dibangun atau persamaan utama akan dikonversi.
                </p>
              </div>
            </div>
          </section>

          {errorStatus && (
            <div className="bg-red-500/10 border border-red-500/20 backdrop-blur-md rounded-2xl p-4 flex items-start gap-3 animate-in fade-in slide-in-from-top-1">
              <AlertCircle className="text-red-400 shrink-0 mt-0.5" size={18} />
              <p className="text-sm text-red-100 leading-relaxed font-medium">
                {errorStatus}
              </p>
            </div>
          )}

          {finalRoot !== null && (
            <div className="bg-teal-500/10 border border-teal-500/20 backdrop-blur-md rounded-2xl p-6 animate-in zoom-in-95 duration-300">
              <div className="flex items-center gap-2 text-teal-400 mb-2">
                <div className="w-2 h-2 bg-teal-400 rounded-full animate-pulse" />
                <span className="text-[10px] font-bold uppercase tracking-widest opacity-70">Akar Ditemukan</span>
              </div>
              <div className="text-3xl font-mono font-medium text-white tracking-tighter">
                {formatFullNumber(finalRoot, parseInt(precision))}
              </div>
              <div className="flex justify-between items-center mt-3">
                <span className="text-[10px] text-slate-400 uppercase tracking-widest font-mono">Iterasi: {results.length - 1}</span>
                <span className="text-[10px] text-teal-400 uppercase tracking-widest font-mono">Precision: {precision} Digits</span>
              </div>
            </div>
          )}

          <div className="bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-5 relative overflow-hidden group">
             <div className="relative z-10">
               <div className="flex items-center gap-2 mb-2">
                 <Info size={14} className="text-indigo-400/60" />
                 <span className="text-[10px] font-bold uppercase tracking-widest text-slate-400">Rumus Secant</span>
               </div>
               <div className="font-mono text-xs leading-relaxed text-slate-300">
                 xₙ₊₁ = xₙ - f(xₙ) · (xₙ - xₙ₋₁) / (f(xₙ) - f(xₙ₋₁))
               </div>
             </div>
             <div className="absolute top-0 right-0 p-4 opacity-5 group-hover:opacity-10 transition-opacity">
                <Calculator size={64} className="text-indigo-400" />
             </div>
          </div>
        </div>

        {/* Content Area */}
        <div className="lg:col-span-8 space-y-6">
          {/* Chart Section */}
          <section className="bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl overflow-hidden shadow-2xl flex flex-col">
            <div className="flex items-center justify-between px-6 py-4 border-b border-white/10 bg-white/5">
              <div className="flex items-center gap-2">
                <ChartIcon size={18} className="text-indigo-400" />
                <h3 className="text-sm font-semibold tracking-wide">Visualisasi Fungsi</h3>
              </div>
              <div className="flex items-center gap-2">
                <div className="w-3 h-3 rounded-full bg-indigo-500/30 border border-indigo-500"></div>
                <span className="text-[10px] text-slate-400 uppercase tracking-widest">Secant Path</span>
              </div>
            </div>
            
            <div className="h-[350px] w-full p-6">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={chartData} margin={{ top: 20, right: 30, left: 0, bottom: 20 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="rgba(255,255,255,0.05)" />
                  <XAxis 
                    dataKey="x" 
                    type="number" 
                    domain={['auto', 'auto']} 
                    fontSize={10} 
                    tick={{ fill: 'rgba(255,255,255,0.4)' }}
                    axisLine={{ stroke: 'rgba(255,255,255,0.1)' }}
                  />
                  <YAxis 
                    fontSize={10} 
                    tick={{ fill: 'rgba(255,255,255,0.4)' }}
                    axisLine={{ stroke: 'rgba(255,255,255,0.1)' }}
                  />
                  <RechartsTooltip 
                    contentStyle={{ 
                      backgroundColor: 'rgba(15, 23, 42, 0.9)', 
                      borderRadius: '12px', 
                      border: '1px solid rgba(255, 255, 255, 0.1)', 
                      backdropFilter: 'blur(8px)',
                      color: '#fff'
                    }}
                    itemStyle={{ color: '#fff' }}
                  />
                  <ReferenceLine y={0} stroke="rgba(255,255,255,0.2)" />
                  <ReferenceLine x={0} stroke="rgba(255,255,255,0.2)" />
                  {finalRoot !== null && (
                    <ReferenceLine x={finalRoot} stroke="#10b981" strokeDasharray="5 5" label={{ value: 'Root', position: 'top', fill: '#10b981', fontSize: 10 }} />
                  )}
                  <Line 
                    type="monotone" 
                    dataKey="y" 
                    stroke="#818cf8" 
                    strokeWidth={2} 
                    dot={false}
                    animationDuration={1000}
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </section>

          {/* Results Table */}
          <section className="bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl overflow-hidden shadow-2xl flex flex-col">
            <div className="flex items-center justify-between px-6 py-4 border-b border-white/10 bg-white/5">
              <div className="flex items-center gap-2">
                <TableIcon size={18} className="text-indigo-400" />
                <h3 className="text-sm font-semibold tracking-wide">Iteration Breakdown</h3>
              </div>
              {results.length > 0 && (
                <button 
                  onClick={exportToExcel}
                  className="flex items-center gap-2 px-3 py-1.5 bg-indigo-500/10 hover:bg-indigo-500/20 border border-indigo-500/30 rounded-lg text-indigo-300 text-[10px] font-bold uppercase tracking-widest transition-all active:scale-95"
                >
                  <Download size={14} />
                  Export .XLSX
                </button>
              )}
            </div>
            
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-white/5 border-b border-white/5">
                    <th className="px-6 py-3 text-[10px] font-bold uppercase tracking-widest text-slate-500">r</th>
                    <th className="px-6 py-3 text-[10px] font-bold uppercase tracking-widest text-slate-500">xr</th>
                    <th className="px-6 py-3 text-[10px] font-bold uppercase tracking-widest text-slate-500">f(xr)</th>
                    <th className="px-6 py-3 text-[10px] font-bold uppercase tracking-widest text-slate-500 text-right">|xr - xr-1|</th>
                  </tr>
                </thead>
                <tbody className="text-sm font-mono divide-y divide-white/5">
                  {results.length > 0 ? (
                    results.map((r, idx) => (
                      <tr 
                        key={idx} 
                        onClick={() => setSelectedIter(r)}
                        className={cn(
                          "bg-transparent hover:bg-white/10 transition-all cursor-pointer group border-l-2 border-transparent",
                          selectedIter?.iteration === r.iteration && "bg-indigo-500/20 shadow-inner",
                          idx === results.length - 1 && finalRoot !== null && selectedIter?.iteration !== r.iteration && "bg-indigo-500/10",
                          r.iteration > 0 && r.error < parseFloat(tolerance) && "bg-teal-500/10 border-l-teal-500/50"
                        )}
                      >
                        <td className="px-6 py-4 text-slate-500 flex items-center gap-2">
                          {r.iteration}
                          {r.iteration > 0 && r.error < parseFloat(tolerance) && (
                            <Check size={12} className="text-teal-400 stroke-[3]" />
                          )}
                        </td>
                        <td className="px-6 py-4 text-slate-100 font-medium">
                          {formatFullNumber(r.x, parseInt(precision))}
                        </td>
                        <td className="px-6 py-4 text-slate-400 italic">
                          {formatFullNumber(r.fx, parseInt(precision))}
                        </td>
                        <td className="px-6 py-4 text-[10px] text-slate-500 text-right">
                          <div className="flex flex-col items-end gap-1">
                            <span className={cn(
                              "font-mono",
                              r.iteration > 0 && r.error < parseFloat(tolerance) ? "text-teal-400 font-bold" : ""
                            )}>
                              {r.error === null ? "-" : formatFullNumber(r.error, Math.max(8, parseInt(precision)))}
                            </span>
                            {r.iteration > 0 && r.error !== null && (
                              <div className="flex items-center gap-1.5 mt-1">
                                <span className={cn(
                                  "px-1.5 py-0.5 rounded border text-[9px] font-bold",
                                  r.error < parseFloat(tolerance) 
                                    ? "text-teal-400 bg-teal-400/10 border-teal-500/20" 
                                    : "text-amber-400 bg-amber-400/10 border-amber-500/20"
                                )}>
                                  {r.error < parseFloat(tolerance) ? "ERR < ε" : "ERR > ε"}
                                </span>
                              </div>
                            )}
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={6} className="px-6 py-20 text-center text-slate-500 italic text-sm">
                        Waiting for calculation signals...
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      </main>

      {/* Footer */}
      <footer className="h-10 bg-indigo-950/40 backdrop-blur-xl border-t border-white/10 px-8 flex items-center justify-between text-[10px] text-slate-500 uppercase tracking-widest z-10">
        <div className="flex gap-6">
          <span>Session: Numerical Analysis</span>
          <span className="hidden sm:inline">Precision: Optimized</span>
        </div>
        <div className="flex gap-4">
          <span className="text-indigo-400">Numerical Engine v1.4.2</span>
          <span className="hidden sm:inline">© 2026 Numérica Lab</span>
        </div>
      </footer>
    </div>
  );
}

