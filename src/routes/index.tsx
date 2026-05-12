import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  Users, Calendar, Briefcase, CheckCircle2, Plus, Trash2,
  Calculator, Printer, FileSpreadsheet, Search, Save, Upload, Moon, Sun,
} from "lucide-react";
import logo from "@/assets/logo.png";

export const Route = createFileRoute("/")({
  component: LeavePage,
  head: () => ({
    meta: [
      { title: "حاسبة رصيد الإجازات السنوية | شركة بلدي للدواجن" },
    ],
  }),
});

type Row = {
  id: number;
  empId: string;
  name: string;
  start: string;
  end: string;
  used: number;
};

const STORAGE_KEY = "baladi-leave-rows-v1";

function calculateLeave(start: string, end: string): number {
  if (!start || !end) return 0;
  const s = new Date(start), e = new Date(end);
  if (isNaN(+s) || isNaN(+e) || e <= s) return 0;
  const days = (+e - +s) / 86400000;
  const years = days / 365.25;
  if (years <= 5) return years * 21;
  return 5 * 21 + (years - 5) * 30;
}

function tenureYears(start: string, end: string): number {
  if (!start || !end) return 0;
  const s = new Date(start), e = new Date(end);
  if (isNaN(+s) || isNaN(+e) || e <= s) return 0;
  return (+e - +s) / 86400000 / 365.25;
}

function LeavePage() {
  const [rows, setRows] = useState<Row[]>([
    { id: 1, empId: "", name: "", start: "", end: "", used: 0 },
  ]);
  const [calculated, setCalculated] = useState(false);
  const [search, setSearch] = useState("");
  const [dark, setDark] = useState(false);
  const [toast, setToast] = useState<{ msg: string; type: "success" | "error" } | null>(null);
  const [nextId, setNextId] = useState(2);

  // Load saved
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed) && parsed.length) {
          setRows(parsed);
          setNextId(Math.max(...parsed.map((r: Row) => r.id)) + 1);
        }
      }
    } catch {}
  }, []);

  useEffect(() => {
    document.documentElement.classList.toggle("dark", dark);
  }, [dark]);

  const showToast = (msg: string, type: "success" | "error" = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 2800);
  };

  const updateRow = (id: number, patch: Partial<Row>) => {
    setRows((prev) => prev.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  };

  const addRow = () => {
    setRows((prev) => [...prev, { id: nextId, empId: "", name: "", start: "", end: "", used: 0 }]);
    setNextId((n) => n + 1);
  };

  const deleteRow = (id: number) => {
    if (rows.length <= 1) return;
    setRows((prev) => prev.filter((r) => r.id !== id));
  };

  const filteredRows = useMemo(() => {
    if (!search.trim()) return rows;
    const q = search.toLowerCase();
    return rows.filter(
      (r) => r.name.toLowerCase().includes(q) || r.empId.toLowerCase().includes(q),
    );
  }, [rows, search]);

  const stats = useMemo(() => {
    let entitled = 0, used = 0, remaining = 0, validCount = 0, tenureSum = 0;
    rows.forEach((r) => {
      const valid = r.start && r.end && new Date(r.end) > new Date(r.start);
      if (valid && calculated) {
        const t = calculateLeave(r.start, r.end);
        entitled += t;
        used += r.used;
        remaining += t - r.used;
        validCount++;
        tenureSum += tenureYears(r.start, r.end);
      }
    });
    return {
      entitled, used, remaining, count: rows.length,
      avgTenure: validCount ? tenureSum / validCount : 0,
    };
  }, [rows, calculated]);

  const handleCalculate = () => {
    const hasValid = rows.some((r) => r.start && r.end && new Date(r.end) > new Date(r.start));
    if (!hasValid) {
      showToast("يرجى إدخال تواريخ صحيحة لموظف واحد على الأقل", "error");
      return;
    }
    setCalculated(true);
    showToast("تم حساب الرصيد بنجاح");
  };

  const handleSave = () => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
    showToast("تم حفظ البيانات محلياً");
  };

  const handleClear = () => {
    if (!confirm("هل تريد مسح جميع البيانات؟")) return;
    localStorage.removeItem(STORAGE_KEY);
    setRows([{ id: 1, empId: "", name: "", start: "", end: "", used: 0 }]);
    setNextId(2);
    setCalculated(false);
    showToast("تم مسح البيانات");
  };

  const exportXLSX = () => {
    const data: (string | number)[][] = [
      ["الرقم الوظيفي", "الاسم", "تاريخ المباشرة", "آخر يوم عمل", "سنوات الخدمة", "المستخدم", "الرصيد المستحق", "الرصيد المتبقي"],
    ];
    rows.forEach((r) => {
      const valid = r.start && r.end && new Date(r.end) > new Date(r.start);
      const total = valid ? calculateLeave(r.start, r.end) : 0;
      const ten = valid ? tenureYears(r.start, r.end) : 0;
      data.push([
        r.empId || "—", r.name || "—", r.start, r.end,
        valid ? Number(ten.toFixed(2)) : "",
        r.used,
        valid ? Number(total.toFixed(2)) : "",
        valid ? Number((total - r.used).toFixed(2)) : "",
      ]);
    });
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [{ wch: 14 }, { wch: 22 }, { wch: 15 }, { wch: 15 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 16 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "رصيد الإجازات");
    XLSX.writeFile(wb, "رصيد_الإجازات_بلدي.xlsx");
    showToast("تم تصدير الملف بنجاح");
  };

  const handlePrint = () => {
    if (!calculated) setCalculated(true);
    setTimeout(() => window.print(), 100);
  };

  return (
    <div dir="rtl" lang="ar" className="min-h-screen pb-12 font-[Tajawal] bg-background text-foreground">
      {/* Header */}
      <header className="header relative overflow-hidden pt-9 pb-16" style={{ background: "var(--gradient-header)", boxShadow: "var(--shadow-header)" }}>
        <div className="absolute inset-0 pointer-events-none" style={{
          backgroundImage:
            "radial-gradient(circle at 15% 25%, oklch(0.85 0.17 90 / 0.12) 0%, transparent 40%), radial-gradient(circle at 85% 75%, oklch(1 0 0 / 0.08) 0%, transparent 45%)",
        }} />
        <div className="absolute top-0 inset-x-0 h-0.5" style={{ background: "linear-gradient(90deg, transparent, var(--gold) 50%, transparent)", opacity: 0.7 }} />
        <div className="relative max-w-[1180px] mx-auto px-6 flex items-center gap-5">
          <div className="w-[88px] h-[88px] flex-shrink-0 rounded-2xl bg-white/10 backdrop-blur-sm flex items-center justify-center ring-1 ring-white/15">
            <img src={logo} alt="شعار شركة بلدي للدواجن" width={72} height={72} className="object-contain" />
          </div>
          <div className="flex-1 min-w-0">
            <h1 className="font-[Cairo] font-extrabold text-white text-[clamp(22px,3.5vw,34px)] m-0 leading-tight tracking-wide" style={{ textShadow: "0 2px 8px rgb(0 0 0 / 0.15)" }}>
              حاسبة رصيد الإجازات السنوية
            </h1>
            <p className="m-0 mt-1.5 text-white/85 text-sm font-medium font-[Cairo]">
              شركة بلدي للدواجن &nbsp;•&nbsp; وفق نظام العمل السعودي{" "}
              <span style={{ color: "var(--gold)" }} className="font-bold">المادة 109</span>
            </p>
          </div>
          <button
            onClick={() => setDark((d) => !d)}
            className="no-print w-10 h-10 rounded-xl bg-white/10 hover:bg-white/20 ring-1 ring-white/15 text-white flex items-center justify-center transition"
            aria-label="تبديل الوضع"
          >
            {dark ? <Sun className="w-5 h-5" /> : <Moon className="w-5 h-5" />}
          </button>
        </div>
      </header>

      <div className="max-w-[1180px] mx-auto px-6 -mt-10 relative">
        {/* Stats */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-7">
          <StatCard color="oklch(0.32 0.13 265)" bg="oklch(0.32 0.13 265 / 0.10)" Icon={Users} label="عدد الموظفين" value={stats.count} />
          <StatCard color="oklch(0.62 0.18 145)" bg="oklch(0.62 0.18 145 / 0.12)" Icon={Calendar} label="إجمالي الرصيد المستحق" value={stats.entitled.toFixed(1)} />
          <StatCard color="oklch(0.6 0.16 60)" bg="oklch(0.85 0.17 90 / 0.18)" Icon={Briefcase} label="إجمالي المستخدم" value={stats.used.toFixed(1)} />
          <StatCard color="oklch(0.45 0.18 260)" bg="oklch(0.45 0.18 260 / 0.10)" Icon={CheckCircle2} label="إجمالي المتبقي" value={stats.remaining.toFixed(1)} />
        </div>

        {/* Average tenure strip */}
        {calculated && stats.avgTenure > 0 && (
          <div className="mb-5 px-5 py-3 rounded-xl bg-card border flex items-center justify-between text-sm" style={{ boxShadow: "var(--shadow-card)" }}>
            <span className="text-muted-foreground">متوسط سنوات الخدمة للموظفين المحسوبين</span>
            <span className="font-[Cairo] font-bold text-base" style={{ color: "var(--accent)" }}>
              {stats.avgTenure.toFixed(2)} سنة
            </span>
          </div>
        )}

        {/* Card */}
        <div className="bg-card rounded-xl border overflow-hidden" style={{ boxShadow: "var(--shadow-card)" }}>
          <div className="px-6 py-5 border-b flex items-center justify-between gap-3 flex-wrap">
            <div>
              <h2 className="m-0 font-[Cairo] font-bold text-lg" style={{ color: "var(--primary)" }}>سجل الموظفين</h2>
              <p className="m-0 mt-1 text-xs text-muted-foreground">21 يوم/سنة لأول 5 سنوات &nbsp;•&nbsp; 30 يوم/سنة بعد 5 سنوات</p>
            </div>
            <div className="flex items-center gap-2 flex-wrap no-print">
              <div className="relative">
                <Search className="w-4 h-4 absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground" />
                <input
                  type="text"
                  placeholder="بحث بالاسم أو الرقم"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  className="pr-9 pl-3 py-2 text-sm border rounded-lg bg-background w-48 focus:outline-none focus:ring-2"
                  style={{ borderColor: "var(--border)" }}
                />
              </div>
              <span className="text-xs px-3.5 py-1.5 rounded-full font-bold border"
                style={{ background: "oklch(0.85 0.17 90 / 0.15)", color: "oklch(0.45 0.12 70)", borderColor: "oklch(0.85 0.17 90 / 0.35)" }}>
                {rows.length} موظف
              </span>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full border-collapse">
              <thead style={{ background: "oklch(0.96 0.012 250)" }}>
                <tr>
                  {["الرقم الوظيفي", "الاسم", "تاريخ المباشرة", "آخر يوم عمل", "المستخدم", "الرصيد المستحق", "الرصيد المتبقي", ""].map((h, i) => (
                    <th key={i} className="px-4 py-3.5 text-right text-[11px] font-bold uppercase tracking-wider text-muted-foreground whitespace-nowrap font-[Cairo]">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredRows.map((row) => {
                  const valid = row.start && row.end && new Date(row.end) > new Date(row.start);
                  const total = valid && calculated ? calculateLeave(row.start, row.end) : null;
                  const remaining = total !== null ? total - row.used : null;
                  return (
                    <tr key={row.id} className="border-t transition hover:bg-secondary/40">
                      <td className="px-4 py-3"><Input value={row.empId} placeholder="EMP-001" onChange={(v) => updateRow(row.id, { empId: v })} /></td>
                      <td className="px-4 py-3"><Input value={row.name} placeholder="اسم الموظف" onChange={(v) => updateRow(row.id, { name: v })} /></td>
                      <td className="px-4 py-3"><Input type="date" value={row.start} onChange={(v) => updateRow(row.id, { start: v })} /></td>
                      <td className="px-4 py-3"><Input type="date" value={row.end} onChange={(v) => updateRow(row.id, { end: v })} /></td>
                      <td className="px-4 py-3">
                        <input
                          type="number" min={0} step={0.5} value={row.used}
                          onChange={(e) => updateRow(row.id, { used: parseFloat(e.target.value) || 0 })}
                          className="w-24 text-center px-2 py-2 border rounded-lg bg-background text-sm tabular-nums focus:outline-none focus:ring-2"
                          style={{ borderColor: "var(--border)" }}
                        />
                      </td>
                      <td className="px-4 py-3">
                        {total !== null
                          ? <span className="font-[Cairo] font-extrabold text-[15px] tabular-nums" style={{ color: "var(--primary)" }}>{total.toFixed(2)}</span>
                          : <span className="text-muted-foreground/50">—</span>}
                      </td>
                      <td className="px-4 py-3">
                        {remaining !== null
                          ? <span className="inline-flex px-3 py-1 rounded-full text-sm font-bold tabular-nums"
                              style={{
                                background: remaining >= 0 ? "var(--success-bg)" : "var(--danger-bg)",
                                color: remaining >= 0 ? "var(--success)" : "var(--danger)",
                              }}>
                              {remaining.toFixed(2)}
                            </span>
                          : <span className="text-muted-foreground/50">—</span>}
                      </td>
                      <td className="px-4 py-3 no-print">
                        <button
                          onClick={() => deleteRow(row.id)}
                          disabled={rows.length === 1}
                          className="w-9 h-9 rounded-lg inline-flex items-center justify-center text-muted-foreground hover:text-destructive hover:bg-destructive/10 disabled:opacity-30 disabled:cursor-not-allowed transition"
                          aria-label="حذف"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </td>
                    </tr>
                  );
                })}
                {filteredRows.length === 0 && (
                  <tr><td colSpan={8} className="text-center py-10 text-muted-foreground text-sm">لا توجد نتائج مطابقة للبحث</td></tr>
                )}
              </tbody>
            </table>
          </div>

          <div className="px-6 py-4 border-t flex flex-wrap gap-3 no-print" style={{ background: "oklch(0.985 0.005 250)" }}>
            <Btn onClick={handleCalculate} variant="primary" Icon={Calculator}>حساب الرصيد</Btn>
            <Btn onClick={addRow} variant="secondary" Icon={Plus}>إضافة موظف</Btn>
            <Btn onClick={handleSave} variant="secondary" Icon={Save}>حفظ</Btn>
            <Btn onClick={handleClear} variant="secondary" Icon={Upload}>مسح الكل</Btn>
            <div className="ms-auto flex flex-wrap gap-3">
              <Btn onClick={handlePrint} variant="outline" Icon={Printer}>طباعة</Btn>
              <Btn onClick={exportXLSX} variant="gold" Icon={FileSpreadsheet}>تصدير إلى Excel</Btn>
            </div>
          </div>
        </div>

        <p className="mt-6 text-center text-xs text-muted-foreground">
          الحساب وفق المادة 109 من نظام العمل السعودي &nbsp;•&nbsp; شركة بلدي للدواجن © 2026
        </p>
      </div>

      {toast && (
        <div
          className="fixed bottom-6 left-1/2 -translate-x-1/2 px-6 py-3 rounded-xl text-white font-bold text-sm font-[Cairo] z-50"
          style={{
            background: toast.type === "success" ? "var(--success)" : "var(--danger)",
            boxShadow: "0 20px 40px rgb(0 0 0 / 0.2)",
          }}
        >
          {toast.type === "success" ? "✓ " : "⚠ "}{toast.msg}
        </div>
      )}

      <style>{`
        @media print {
          @page { size: A4 landscape; margin: 1cm; }
          .no-print { display: none !important; }
          .header { background: white !important; box-shadow: none !important; padding: 0 0 12px !important; border-bottom: 3px solid var(--gold); }
          .header h1 { color: var(--primary) !important; text-shadow: none !important; }
          .header p { color: var(--muted-foreground) !important; }
        }
      `}</style>
    </div>
  );
}

function Input({ value, onChange, placeholder, type = "text" }: {
  value: string; onChange: (v: string) => void; placeholder?: string; type?: string;
}) {
  return (
    <input
      type={type}
      value={value}
      placeholder={placeholder}
      onChange={(e) => onChange(e.target.value)}
      className="w-full px-3 py-2 border rounded-lg bg-background text-sm focus:outline-none focus:ring-2 transition"
      style={{ borderColor: "var(--border)", direction: type === "date" ? "ltr" : undefined, textAlign: type === "date" ? "right" : undefined }}
    />
  );
}

function StatCard({ Icon, label, value, color, bg }: {
  Icon: React.ComponentType<{ className?: string; style?: React.CSSProperties }>;
  label: string; value: string | number; color: string; bg: string;
}) {
  return (
    <div className="relative bg-card rounded-xl p-5 border overflow-hidden transition hover:-translate-y-0.5" style={{ boxShadow: "var(--shadow-card)" }}>
      <div className="absolute top-0 right-0 w-16 h-16 rounded-full opacity-20" style={{ background: color, transform: "translate(20px,-20px)" }} />
      <div className="w-11 h-11 rounded-lg flex items-center justify-center mb-3.5" style={{ background: bg }}>
        <Icon className="w-5 h-5" style={{ color }} />
      </div>
      <div className="text-[26px] font-extrabold tabular-nums font-[Cairo]" style={{ color: "var(--primary)" }}>{value}</div>
      <div className="text-xs font-semibold text-muted-foreground mt-0.5">{label}</div>
    </div>
  );
}

function Btn({ children, onClick, variant, Icon }: {
  children: React.ReactNode; onClick: () => void; variant: "primary" | "secondary" | "gold" | "outline";
  Icon: React.ComponentType<{ className?: string }>;
}) {
  const styles: Record<string, React.CSSProperties> = {
    primary: { background: "var(--gradient-header)", color: "white", boxShadow: "0 2px 8px oklch(0.32 0.13 265 / 0.3)" },
    secondary: { background: "white", color: "var(--primary)", border: "1.5px solid var(--border)" },
    gold: { background: "var(--gradient-gold)", color: "oklch(0.25 0.08 60)", boxShadow: "0 2px 8px oklch(0.85 0.17 90 / 0.45)" },
    outline: { background: "white", color: "var(--primary)", border: "1.5px solid var(--primary)" },
  };
  return (
    <button
      onClick={onClick}
      className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl font-[Cairo] text-sm font-bold tracking-wide transition hover:-translate-y-0.5"
      style={styles[variant]}
    >
      <Icon className="w-4 h-4" />
      {children}
    </button>
  );
}
