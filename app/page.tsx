import { useEffect, useMemo, useRef, useState } from "react";
import {
  Archive,
  ArrowLeft,
  BarChart3,
  Bell,
  Boxes,
  CalendarDays,
  Check,
  CheckCircle2,
  ChevronRight,
  CircleAlert,
  Clock3,
  Database,
  Download,
  FileSpreadsheet,
  FileText,
  HardDrive,
  Keyboard,
  LayoutDashboard,
  ListFilter,
  Lock,
  Menu,
  MoreHorizontal,
  Moon,
  Plus,
  RefreshCw,
  RotateCcw,
  ScanLine,
  Search,
  Settings,
  ShieldCheck,
  Sun,
  Table2,
  Trash2,
  Upload,
  UserSearch,
  Users,
  X,
  Zap,
} from "lucide-react";
import { DEMO_EMPLOYEES, formatDate, normalizeHeader, normalizeMatricula, readHistory, todayIso, validDate } from "./lib/data";
import type { Employee, RecordRow, CompletedBox, ImportResult } from "./lib/data";

type ScanStatus = "idle" | "capturing" | "reading" | "consulting" | "success";

type BaseInfo = {
  fileName: string;
  sheetName: string;
  count: number;
  importedAt: string;
  duplicates: number;
  conflicts: number;
  invalidRows: number;
  sheetsFound: number;
  demo?: boolean;
};

type AppView = "dashboard" | "batch" | "boxes" | "reports" | "employees" | "settings";
type Theme = "light" | "dark";

const HISTORY_KEY = "arquivo-agil-autoral-box-history-v1";
const THEME_KEY = "arquivo-agil-autoral-theme";
const DUPLICATE_KEY = "arquivo-agil-autoral-prevent-duplicates";

function createLocalId() {
  return typeof globalThis.crypto?.randomUUID === "function"
    ? globalThis.crypto.randomUUID()
    : `caixa-${Date.now()}-${Math.random().toString(36).slice(2, 12)}`;
}

async function downloadWorkbook(workbook: import("exceljs").Workbook, filename: string) {
  const buffer = await workbook.xlsx.writeBuffer();
  const url = URL.createObjectURL(new Blob([new Uint8Array(buffer)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

export default function Home() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [mobileNav, setMobileNav] = useState(false);
  const [view, setView] = useState<AppView>("dashboard");
  const [theme, setTheme] = useState<Theme>("light");
  const [step, setStep] = useState<1 | 2 | 3>(1);
  const [label, setLabel] = useState("");
  const [date, setDate] = useState(todayIso);
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [baseInfo, setBaseInfo] = useState<BaseInfo | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const [records, setRecords] = useState<RecordRow[]>([]);
  const [scanStatus, setScanStatus] = useState<ScanStatus>("idle");
  const [manualInput, setManualInput] = useState("");
  const [lastEmployee, setLastEmployee] = useState<Employee | null>(null);
  const [toast, setToast] = useState<string | null>(null);
  const [showSuccess, setShowSuccess] = useState(false);
  const [allHistory, setAllHistory] = useState<CompletedBox[]>([]);
  const [historyReady, setHistoryReady] = useState(false);
  const [preventDuplicates, setPreventDuplicates] = useState(true);
  const [storageAvailable, setStorageAvailable] = useState(true);
  const [isExporting, setIsExporting] = useState(false);
  const [exportedFingerprint, setExportedFingerprint] = useState("");
  const [conflicts, setConflicts] = useState<Set<string>>(new Set());
  const [importSeconds, setImportSeconds] = useState(0);
  const importWorker = useRef<Worker | null>(null);
  const operationLock = useRef(false);
  const recordSequence = useRef(0);
  const boxId = useRef("");
  const toastTimer = useRef<number | null>(null);
  const employeeIndex = useMemo(() => new Map(employees.map((employee) => [employee.matricula, employee])), [employees]);
  const fingerprint = JSON.stringify([label.trim(), date, records]);
  const hasUnexportedRecords = records.length > 0 && fingerprint !== exportedFingerprint;
  const history = useMemo(() => allHistory.filter((box) => !!box.demo === !!baseInfo?.demo), [allHistory, baseInfo?.demo]);

  function savePreference(key: string, value: string) {
    try { window.localStorage.setItem(key, value); }
    catch { setStorageAvailable(false); }
  }

  useEffect(() => {
    try {
      const nextTheme: Theme = window.localStorage.getItem(THEME_KEY) === "dark" ? "dark" : "light";
      document.documentElement.dataset.theme = nextTheme;
      setTheme(nextTheme);
      setAllHistory(readHistory(window.localStorage.getItem(HISTORY_KEY)));
      setPreventDuplicates(window.localStorage.getItem(DUPLICATE_KEY) !== "false");
    } catch { setStorageAvailable(false); }
    setHistoryReady(true);
    return () => { importWorker.current?.terminate(); if (toastTimer.current) clearTimeout(toastTimer.current); };
  }, []);

  useEffect(() => {
    if (historyReady) savePreference(HISTORY_KEY, JSON.stringify(allHistory.filter((box) => !box.demo).slice(0, 1000)));
  }, [allHistory, historyReady]);

  useEffect(() => {
    if (!hasUnexportedRecords && !isImporting && !isExporting) return;
    const warn = (event: BeforeUnloadEvent) => { event.preventDefault(); event.returnValue = ""; };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [hasUnexportedRecords, isImporting, isExporting]);

  useEffect(() => {
    if (!isImporting) return;
    const timer = window.setInterval(() => setImportSeconds((seconds) => seconds + 1), 1000);
    return () => clearInterval(timer);
  }, [isImporting]);

  function changeView(nextView: AppView) {
    setView(nextView);
    setMobileNav(false);
    setShowSuccess(false);
  }

  function toggleTheme() {
    const nextTheme: Theme = theme === "light" ? "dark" : "light";
    setTheme(nextTheme);
    document.documentElement.dataset.theme = nextTheme;
    savePreference(THEME_KEY, nextTheme);
  }

  function updatePreventDuplicates(value: boolean) {
    setPreventDuplicates(value);
    savePreference(DUPLICATE_KEY, String(value));
  }

  const scanCopy: Record<ScanStatus, { title: string; detail: string }> = {
    idle: { title: "Pronto para consultar", detail: "Digite a matrícula real no campo abaixo" },
    capturing: { title: "Recebendo matrícula", detail: "Normalizando o número informado" },
    reading: { title: "Matrícula validada", detail: "Preparando a consulta na base carregada" },
    consulting: { title: "Consultando a base do SOC", detail: "Buscando o funcionário localmente" },
    success: { title: "Funcionário identificado", detail: "Registro incluído na caixa. Digite a próxima matrícula." },
  };

  function flash(message: string) {
    setToast(message);
    if (toastTimer.current) clearTimeout(toastTimer.current);
    toastTimer.current = window.setTimeout(() => setToast(null), 6500);
  }

  function requestImport() {
    if (operationLock.current) return;
    if (records.length) { flash("Conclua ou reinicie a caixa atual antes de substituir a base."); return; }
    fileInputRef.current?.click();
  }

  function openBatch() {
    if (operationLock.current) return;
    if (!baseInfo) { flash("Importe a planilha do SOC antes de iniciar uma caixa."); requestImport(); return; }
    changeView("batch");
  }

  function resetSession(force = false) {
    if (operationLock.current) return false;
    if (!force && hasUnexportedRecords && !window.confirm("Descartar os registros ainda não exportados desta caixa?")) return false;
    changeView("dashboard");
    setStep(1); setLabel(""); setDate(todayIso()); setRecords([]);
    setScanStatus("idle"); setLastEmployee(null); setManualInput("");
    setShowSuccess(false); setExportedFingerprint(""); boxId.current = "";
    return true;
  }

  function resetBase() {
    if (!resetSession()) return;
    setEmployees([]); setConflicts(new Set()); setBaseInfo(null);
    flash("Base removida da memória desta página.");
  }

  function loadDemo() {
    if (operationLock.current) return;
    if (records.length) { flash("Conclua ou reinicie a caixa atual antes de iniciar a demonstração."); return; }
    setEmployees(DEMO_EMPLOYEES); setConflicts(new Set());
    setBaseInfo({ fileName: "Exemplo com 8 pessoas fictícias", sheetName: "Demonstração", count: DEMO_EMPLOYEES.length,
      importedAt: new Date().toLocaleString("pt-BR"), duplicates: 0, conflicts: 0, invalidRows: 0, sheetsFound: 1, demo: true });
    setRecords([]); setLastEmployee(null); setManualInput(""); setExportedFingerprint("");
    setLabel("DEMO-001"); setDate(todayIso()); boxId.current = ""; setStep(1); changeView("batch");
    flash("Demonstração ativada. Consulte DEMO001, DEMO002 e DEMO003.");
  }

  function cancelImport() {
    importWorker.current?.terminate(); importWorker.current = null;
    operationLock.current = false; setIsImporting(false);
    flash("Importação cancelada. A base anterior foi mantida.");
  }

  async function importSocBase(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file || operationLock.current) return;
    if (records.length) { flash("Conclua a caixa atual antes de substituir a base."); return; }
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      flash("Use um arquivo .xlsx. No Excel, salve uma cópia nesse formato se a exportação estiver em .xls."); return;
    }
    if (file.size > 50 * 1024 * 1024) { flash("Este arquivo excede 50 MB. Salve uma cópia .xlsx apenas com matrícula e nome."); return; }
    operationLock.current = true; setIsImporting(true); setImportSeconds(0);
    try {
      const worker = new Worker(new URL("./soc-import.worker.ts", import.meta.url), { type: "module" });
      importWorker.current = worker;
      const finish = () => { worker.terminate(); importWorker.current = null; operationLock.current = false; setIsImporting(false); };
      worker.onmessage = (message: MessageEvent<{ result?: ImportResult; error?: string }>) => {
        if (importWorker.current !== worker) return;
        const result = message.data.result;
        if (!result) { finish(); flash(message.data.error ?? "Não foi possível ler a planilha."); return; }
        setEmployees(result.employees); setConflicts(new Set(result.conflicts));
        setBaseInfo({ fileName: file.name, sheetName: result.sheetName, count: result.employees.length,
          importedAt: new Date().toLocaleString("pt-BR"), duplicates: result.duplicates, conflicts: result.conflicts.length,
          invalidRows: result.invalidRows, sheetsFound: result.sheetsFound });
        setRecords([]); setLabel(""); setDate(todayIso()); setStep(1); setLastEmployee(null);
        setManualInput(""); setExportedFingerprint(""); boxId.current = "";
        finish();
        flash(result.employees.length.toLocaleString("pt-BR") + " matrículas disponíveis. Veja o resumo em Base de funcionários.");
      };
      worker.onerror = () => { if (importWorker.current !== worker) return; finish(); flash("Não foi possível processar a planilha. Salve uma cópia .xlsx sem senha e tente novamente."); };
      const buffer = await file.arrayBuffer();
      if (importWorker.current === worker) worker.postMessage(buffer, [buffer]);
    } catch {
      importWorker.current?.terminate(); importWorker.current = null; operationLock.current = false; setIsImporting(false);
      flash("Não foi possível abrir a planilha neste navegador. Tente novamente com um arquivo .xlsx sem senha.");
    }
  }

  function searchManual() {
    if (operationLock.current) return;
    if (!baseInfo) { flash("Importe a planilha do SOC antes de consultar matrículas."); return; }
    if (!label.trim() || !validDate(date)) { setStep(1); flash("Preencha a etiqueta e uma data válida."); return; }
    const matricula = normalizeMatricula(manualInput);
    if (!matricula) { flash("Digite uma matrícula para continuar."); return; }
    if (conflicts.has(matricula)) { flash("Esta matrícula aparece com nomes diferentes na base. Confira o cadastro de origem antes de incluí-la."); return; }
    const employee = employeeIndex.get(matricula);
    if (!employee) { flash("Matrícula não encontrada na base carregada. Confira o número, incluindo os zeros à esquerda."); return; }
    if (preventDuplicates && records.some((row) => row.matricula === matricula)) { flash("Esta matrícula já foi adicionada à caixa."); return; }
    if (!boxId.current) boxId.current = createLocalId();
    setRecords((current) => [...current, { ...employee, id: ++recordSequence.current, etiqueta: label.trim(),
      setor: "Saúde Ocupacional", data: date, status: "Confirmado" }]);
    setLastEmployee(employee); setManualInput(""); setScanStatus("success");
  }

  function updateLabel(value: string) {
    if (operationLock.current) return;
    setLabel(value);
    setRecords((current) => current.map((row) => ({ ...row, etiqueta: value.trim() })));
  }

  function updateDate(value: string) {
    if (operationLock.current) return;
    setDate(value);
    setRecords((current) => current.map((row) => ({ ...row, data: value })));
  }

  function removeRecord(id: number) {
    if (operationLock.current) return;
    const remaining = records.filter((row) => row.id !== id);
    setRecords(remaining); setLastEmployee(remaining.at(-1) ?? null);
    flash("Registro removido da caixa.");
  }

  async function exportWorkbook() {
    if (operationLock.current) return;
    if (!records.length || !label.trim() || !validDate(date)) { flash("Confira os dados da caixa antes de exportar."); return; }
    operationLock.current = true; setIsExporting(true);
    try {
      const { createBoxWorkbook } = await import("./lib/workbooks");
      const workbook = createBoxWorkbook(records, label, date);
      const filename = label.trim().replace(/[^a-zA-Z0-9-]/g, "_").slice(0, 80);
      await downloadWorkbook(workbook, (baseInfo?.demo ? "DEMONSTRACAO_" : "") + "Planilha_" + filename + ".xlsx");
      const summary: CompletedBox = { id: boxId.current || createLocalId(), etiqueta: label.trim(), data: date,
        recordCount: records.length, exportedAt: new Date().toISOString(), demo: !!baseInfo?.demo };
      boxId.current = summary.id;
      setAllHistory((current) => [summary, ...current.filter((item) => item.id !== summary.id)].slice(0, 1000));
      setExportedFingerprint(fingerprint); setShowSuccess(true);
    } catch { flash("Não foi possível gerar o Excel. Os registros continuam disponíveis; tente exportar novamente."); }
    finally { operationLock.current = false; setIsExporting(false); }
  }

  async function exportHistoryReport(selected: CompletedBox[] = history) {
    if (operationLock.current) return;
    if (!selected.length) { flash("Não há caixas nesse período para exportar."); return; }
    operationLock.current = true; setIsExporting(true);
    try {
      const { createHistoryWorkbook } = await import("./lib/workbooks");
      await downloadWorkbook(createHistoryWorkbook(selected), (baseInfo?.demo ? "DEMONSTRACAO_" : "") + "Relatorio_Arquivo_Agil_" + todayIso() + ".xlsx");
      flash("Relatório gerado. Confira o arquivo na pasta de downloads.");
    } catch { flash("Não foi possível gerar o relatório. Tente novamente."); }
    finally { operationLock.current = false; setIsExporting(false); }
  }

  function removeHistoryItem(id: string) {
    if (!window.confirm("Remover apenas este resumo do histórico? O Excel já baixado será mantido.")) return;
    setAllHistory((current) => current.filter((item) => item.id !== id));
    flash("Resumo da caixa removido do histórico local.");
  }

  function clearHistory() {
    if (!window.confirm("Remover todo o histórico local de caixas processadas?")) return;
    setAllHistory((current) => current.filter((item) => !!item.demo !== !!baseInfo?.demo));
    flash("Histórico local removido.");
  }

  const sidebar = (
    <aside className={`sidebar ${mobileNav ? "sidebar-open" : ""}`}>
      <div className="brand">
        <div className="author-lockup">
          <span className="author-symbol"><ScanLine size={19} strokeWidth={2.4} /></span>
          <div className="product-name"><strong>arquivo<span>ágil</span></strong><small>automação documental</small></div>
        </div>
      </div>

      <nav className="main-nav" aria-label="Navegação principal">
        <p className="nav-label">MENU</p>
        <button className={view === "dashboard" ? "nav-item active" : "nav-item"} onClick={() => changeView("dashboard")}>
          <LayoutDashboard size={19} /><span>Visão geral</span>
        </button>
        <button className={view === "batch" ? "nav-item active" : "nav-item"} onClick={openBatch}>
          <ScanLine size={19} /><span>Planilhar documentos</span><span className="nav-dot" />
        </button>
        <button className={view === "boxes" ? "nav-item active" : "nav-item"} onClick={() => changeView("boxes")}>
          <Boxes size={19} /><span>Caixas processadas</span>
        </button>
        <button className={view === "reports" ? "nav-item active" : "nav-item"} onClick={() => changeView("reports")}>
          <FileSpreadsheet size={19} /><span>Relatórios</span>
        </button>
        <p className="nav-label nav-section">SISTEMA</p>
        <button className={view === "settings" ? "nav-item active" : "nav-item"} onClick={() => changeView("settings")}>
          <Settings size={19} /><span>Configurações</span>
        </button>
      </nav>

      <button className={`integration-card ${view === "employees" ? "active" : ""}`} onClick={() => changeView("employees")}>
        <div className="integration-icon"><Database size={18} /></div>
        <div><span>Base de funcionários</span><strong><i className={baseInfo ? "" : "offline-dot"} /> {baseInfo?.demo ? "Demonstração ativa" : baseInfo ? "Base carregada" : "Aguardando importação"}</strong></div>
      </button>
      <div className="profile">
        <div className="avatar">J</div>
        <div><strong>Projeto por Jean</strong><span>Produto • UX • Tecnologia</span></div>
        <MoreHorizontal size={18} />
      </div>
    </aside>
  );

  return (
    <div className="app-shell">
      <input ref={fileInputRef} className="sr-only" type="file" accept=".xlsx" onChange={importSocBase} aria-label="Selecionar planilha do SOC" />
      <a href="#main-content" className="skip-link">Ir para o conteúdo</a>
      {sidebar}
      {mobileNav && <button className="nav-scrim" aria-label="Fechar menu" onClick={() => setMobileNav(false)} />}

      <div className="content-shell">
        <header className="topbar">
          <button className="mobile-menu" aria-label="Abrir menu" onClick={() => setMobileNav(true)}><Menu size={22} /></button>
          <div className="prototype-badge"><span /> ARQUIVO ÁGIL • POR JEAN</div>
          <div className="top-actions">
            <div className="sync-status"><ShieldCheck size={16} /> {baseInfo ? `${baseInfo.count.toLocaleString("pt-BR")} funcionários • dados locais` : "Nenhuma base carregada"}</div>
            <button className="theme-toggle" aria-label={theme === "light" ? "Ativar modo escuro" : "Ativar modo claro"} title={theme === "light" ? "Modo escuro" : "Modo claro"} onClick={toggleTheme}>
              {theme === "light" ? <Moon size={18} /> : <Sun size={18} />}
              <span>{theme === "light" ? "Escuro" : "Claro"}</span>
            </button>
            <button className="icon-button" aria-label="Status da sessão" onClick={() => flash(baseInfo ? "Base carregada nesta sessão. Exporte a caixa antes de recarregar a página. A conexão automática com o SOC não está ativa." : "Importe a planilha ou inicie a demonstração na visão geral.")}><Bell size={19} /></button>
          </div>
        </header>

        <main className="main-content" id="main-content">
          {baseInfo?.demo && <div className="notice-banner demo-banner" role="status"><Zap size={20} /><div><strong>Demonstração • dados fictícios</strong><span>Use DEMO001, DEMO002 e DEMO003. Os indicadores desta demonstração ficam separados do seu histórico real.</span></div><button onClick={resetBase}>Sair da demonstração</button></div>}
          {!storageAvailable && <div className="notice-banner" role="status"><CircleAlert size={20} /><div><strong>Armazenamento do navegador indisponível</strong><span>Histórico e preferências serão mantidos somente enquanto esta página estiver aberta.</span></div></div>}
          {!!label && view !== "batch" && <div className="notice-banner"><Archive size={20} /><div><strong>Caixa {label} em andamento</strong><span>{records.length} registros nesta sessão. Retome para revisar ou exportar.</span></div><button onClick={openBatch}>Continuar caixa</button></div>}
          {view === "dashboard" ? (
            <Dashboard
              onStart={openBatch}
              hasActiveBox={!!label || records.length > 0}
              onDemo={loadDemo}
              onToast={flash}
              baseInfo={baseInfo}
              isImporting={isImporting}
              onImport={requestImport}
              onResetBase={resetBase}
              history={history}
            />
          ) : view === "batch" ? (
            <BatchFlow
              step={step}
              setStep={setStep}
              label={label}
              setLabel={updateLabel}
              date={date}
              setDate={updateDate}
              records={records}
              scanStatus={scanStatus}
              scanCopy={scanCopy}
              lastEmployee={lastEmployee}
              manualInput={manualInput}
              setManualInput={(value) => { setManualInput(value); setScanStatus("idle"); }}
              onManual={searchManual}
              onRemove={removeRecord}
              onBack={() => setView("dashboard")}
              onExport={exportWorkbook}
              isExporting={isExporting}
              onReset={() => resetSession()}
              baseInfo={baseInfo}
            />
          ) : view === "boxes" ? (
            <BoxesPage history={history} onRemove={removeHistoryItem} onStart={openBatch} />
          ) : view === "reports" ? (
            <ReportsPage history={history} onExport={exportHistoryReport} isExporting={isExporting} />
          ) : view === "employees" ? (
            <EmployeesPage employees={employees} baseInfo={baseInfo} isImporting={isImporting} onImport={requestImport} onReset={resetBase} />
          ) : (
            <SettingsPage
              theme={theme}
              onToggleTheme={toggleTheme}
              preventDuplicates={preventDuplicates}
              onPreventDuplicates={updatePreventDuplicates}
              baseInfo={baseInfo}
              historyCount={history.length}
              onImport={requestImport}
              onResetBase={resetBase}
              onClearHistory={clearHistory}
            />
          )}
        </main>
      </div>

      {isImporting && <div className="import-overlay" role="status" aria-live="polite"><section><RefreshCw size={30} className="spin" /><h2>Lendo a planilha</h2><p>Processando os registros neste navegador. Bases grandes podem levar alguns instantes.</p><span>{importSeconds}s decorridos • você pode cancelar</span><button className="text-button" onClick={cancelImport}>Cancelar importação</button></section></div>}
      {toast && <div className="toast" role="status"><CircleAlert size={18} />{toast}<button onClick={() => setToast(null)} aria-label="Fechar"><X size={16} /></button></div>}
      {showSuccess && (
        <div className="modal-backdrop" role="presentation">
          <section className="success-modal" role="dialog" aria-modal="true" aria-labelledby="success-title">
            <div className="success-orbit"><Check size={30} /></div>
            <span className="eyebrow">EXPORTAÇÃO CONCLUÍDA</span>
            <h2 id="success-title">Excel gerado com sucesso</h2>
            <p>O arquivo da caixa <strong>{label}</strong> foi gerado. Confira o download no navegador: {records.length} {records.length === 1 ? "registro confirmado" : "registros confirmados"}.</p>
            <div className="file-result"><FileSpreadsheet size={25} /><div><strong>{baseInfo?.demo ? "DEMONSTRACAO_" : ""}Planilha_{label.trim().replace(/[^a-zA-Z0-9-]/g, "_")}.xlsx</strong><span>{records.length} {records.length === 1 ? "linha" : "linhas"} • tabela formatada • filtros ativos</span></div><CheckCircle2 size={20} /></div>
            <button className="primary-button modal-button" onClick={() => resetSession()}>Concluir caixa <ChevronRight size={18} /></button>
            <button className="text-button" onClick={() => setShowSuccess(false)}>Voltar para a revisão</button>
          </section>
        </div>
      )}
    </div>
  );
}

function Dashboard({
  onStart,
  onDemo,
  hasActiveBox,
  onToast,
  baseInfo,
  isImporting,
  onImport,
  onResetBase,
  history,
}: {
  onStart: () => void;
  onDemo: () => void;
  hasActiveBox: boolean;
  onToast: (message: string) => void;
  baseInfo: BaseInfo | null;
  isImporting: boolean;
  onImport: () => void;
  onResetBase: () => void;
  history: CompletedBox[];
}) {
  const totalDocuments = history.reduce((sum, item) => sum + item.recordCount, 0);
  const latestBox = history[0];

  return (
    <div className="page-enter dashboard-page">
      <section className="welcome-row">
        <div>
          <span className="eyebrow">AUTOMAÇÃO DOCUMENTAL • EXCEL</span>
          <h1>Documentos organizados,<br /><em>sem trabalho repetitivo.</em></h1>
          <p>Importe a base oficial, consulte matrículas reais e gere a planilha de cada caixa em um único fluxo.</p>
        </div>
        <button className="primary-button hero-button" onClick={onStart}><Plus size={19} /> {hasActiveBox ? "Continuar caixa" : baseInfo ? "Planilhar nova caixa" : "Importar base para começar"}</button>
      </section>

      <section className="stats-grid" aria-label="Indicadores da operação">
        <StatCard icon={<Users size={20} />} tone="navy" value={baseInfo ? baseInfo.count.toLocaleString("pt-BR") : "0"} label="Funcionários na base" note={baseInfo ? "Dados da planilha importada" : "Importe a base do SOC"} />
        <StatCard icon={<FileText size={20} />} tone="teal" value={totalDocuments.toLocaleString("pt-BR")} label="Documentos exportados" note={history.length ? `${history.length} caixas no histórico local` : "Nenhuma caixa exportada ainda"} />
        <StatCard icon={<Boxes size={20} />} tone="amber" value={history.length.toLocaleString("pt-BR")} label="Caixas processadas" note={latestBox ? `Última: ${latestBox.etiqueta}` : "Histórico iniciado na primeira exportação"} />
        <StatCard icon={<Lock size={20} />} tone="violet" value="Local" label="Processamento dos dados" note="Sem envio da base ao servidor" />
      </section>

      <section className={`soc-base-card ${baseInfo ? "loaded" : ""}`} aria-label="Base de funcionários do SOC">
        <div className="soc-base-main">
          <div className="soc-base-icon">{baseInfo ? <CheckCircle2 size={25} /> : <Database size={25} />}</div>
          <div>
            <span className="eyebrow">BASE DE FUNCIONÁRIOS</span>
            <h2>{baseInfo?.demo ? "Base fictícia pronta para demonstração" : baseInfo ? "Planilha do SOC pronta para consulta" : "Importe a exportação real do SOC"}</h2>
            <p>{baseInfo ? baseInfo.fileName : "O arquivo precisa ter as colunas de matrícula e nome do funcionário."}</p>
          </div>
        </div>
        <div className="soc-base-metrics">
          <div><span>Funcionários</span><strong>{baseInfo ? baseInfo.count.toLocaleString("pt-BR") : "0"}</strong></div>
          <div><span>Aba utilizada</span><strong>{baseInfo?.sheetName ?? "—"}</strong></div>
          <div><span>Importada em</span><strong>{baseInfo?.importedAt ?? "—"}</strong></div>
        </div>
        <div className="soc-base-actions">
          <button className="base-import-button" onClick={onImport} disabled={isImporting}>
            {isImporting ? <RefreshCw className="spin" size={17} /> : <Upload size={17} />}
            {isImporting ? "Lendo planilha..." : baseInfo ? "Substituir base" : "Importar base real do SOC"}
          </button>
          {baseInfo && <button className="base-reset-button" onClick={onResetBase}>Remover base</button>}
          <span className="local-privacy"><Lock size={14} /> Processamento local: o arquivo não é enviado ao servidor.</span>
        </div>
      </section>

      {baseInfo && <ImportSummary baseInfo={baseInfo} />}
      {!baseInfo && <section className="presentation-card"><div><span className="eyebrow">CONHEÇA O FLUXO</span><h2>Pronto para apresentar</h2><p>Experimente com 8 pessoas fictícias: configure uma caixa, consulte três matrículas, confira o Excel e abra os relatórios.</p></div><button className="base-import-button" onClick={onDemo} disabled={isImporting}><Zap size={18} /> Iniciar demonstração</button></section>}
      <section className="dashboard-grid">
        <article className="start-card">
          <div className="start-copy">
            <div className="kicker"><Zap size={15} fill="currentColor" /> FLUXO AUTOMATIZADO</div>
            <h2>{hasActiveBox ? "Continue de onde parou" : "Comece uma nova caixa"}</h2>
            <p>Você informa a etiqueta uma vez. O sistema cuida do restante.</p>
            <div className="flow-chips">
              <span><b>01</b> Configurar</span><ChevronRight size={16} />
              <span><b>02</b> Consultar</span><ChevronRight size={16} />
              <span><b>03</b> Exportar</span>
            </div>
            <button className="secondary-button" onClick={onStart}>Abrir planilhamento <ChevronRight size={17} /></button>
          </div>
          <div className="scan-illustration" aria-hidden="true">
            <div className="document-stack back" />
            <div className="document-stack middle" />
            <div className="mini-document">
              <div className="doc-logo"><Archive size={15} /></div>
              <span className="doc-line long" /><span className="doc-line" /><span className="doc-line short" />
              <div className="mini-matricula"><small>MATRÍCULA</small><strong>REAL</strong></div>
              <div className="mini-scan-line" />
            </div>
            <div className="check-float"><Check size={17} /></div>
          </div>
        </article>

        <article className="recent-card">
          <div className="card-heading"><div><span className="eyebrow">{history.length ? "HISTÓRICO LOCAL" : "PRIMEIRO TESTE"}</span><h2>{history.length ? "Caixas recentes" : "Validação controlada"}</h2></div><button aria-label="Ajuda" onClick={() => onToast(history.length ? "O histórico guarda apenas etiqueta, data e quantidade de documentos." : "Comece com poucas fichas e confira cada linha antes de exportar.")}><MoreHorizontal size={20} /></button></div>
          <div className="recent-list">
            {history.length ? history.slice(0, 3).map((item) => (
              <div className="recent-item" key={item.id}><div className="box-icon"><CheckCircle2 size={19} /></div><div className="recent-main"><strong>{item.etiqueta}</strong><span>{item.recordCount} {item.recordCount === 1 ? "documento" : "documentos"} • {formatDate(item.data)}</span></div><span className="status-pill"><i />Exportada</span></div>
            )) : <>
              <div className="recent-item"><div className="box-icon"><Upload size={19} /></div><div className="recent-main"><strong>1. Importe a base do SOC</strong><span>Use a exportação .xlsx atualizada.</span></div></div>
              <div className="recent-item"><div className="box-icon"><Search size={19} /></div><div className="recent-main"><strong>2. Teste de 3 a 5 matrículas</strong><span>Compare os nomes encontrados com o SOC.</span></div></div>
              <div className="recent-item"><div className="box-icon"><FileSpreadsheet size={19} /></div><div className="recent-main"><strong>3. Confira o Excel gerado</strong><span>Valide etiqueta, nome, matrícula, setor e data de temporalidade.</span></div></div>
            </>}
          </div>
        </article>
      </section>
    </div>
  );
}

function PageHeader({ eyebrow, title, description, action }: { eyebrow: string; title: string; description: string; action?: React.ReactNode }) {
  return (
    <section className="section-page-header">
      <div><span className="eyebrow">{eyebrow}</span><h1>{title}</h1><p>{description}</p></div>
      {action}
    </section>
  );
}

function BoxesPage({ history, onRemove, onStart }: { history: CompletedBox[]; onRemove: (id: string) => void; onStart: () => void }) {
  const [query, setQuery] = useState("");
  const filtered = useMemo(() => history.filter((item) => item.etiqueta.toLowerCase().includes(query.toLowerCase())), [history, query]);
  const totalDocuments = history.reduce((sum, item) => sum + item.recordCount, 0);

  return (
    <div className="page-enter operations-page">
      <PageHeader eyebrow="CONTROLE DOCUMENTAL" title="Caixas processadas" description="Acompanhe as exportações realizadas neste navegador sem armazenar nomes ou matrículas." action={<button className="primary-button" onClick={onStart}><Plus size={18} /> Nova caixa</button>} />
      <section className="compact-stats">
        <StatCard icon={<Boxes size={20} />} tone="navy" value={history.length.toLocaleString("pt-BR")} label="Caixas exportadas" note="Histórico local deste navegador" />
        <StatCard icon={<FileText size={20} />} tone="teal" value={totalDocuments.toLocaleString("pt-BR")} label="Documentos planilhados" note="Soma das caixas exportadas" />
        <StatCard icon={<CalendarDays size={20} />} tone="amber" value={history[0] ? formatDate(history[0].data) : "—"} label="Última data de temporalidade" note={history[0]?.etiqueta ?? "Nenhuma exportação registrada"} />
      </section>
      <section className="data-panel">
        <div className="panel-toolbar">
          <div><span className="eyebrow">HISTÓRICO</span><h2>Exportações concluídas</h2></div>
          <div className="panel-search"><Search size={17} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Buscar por etiqueta" /></div>
        </div>
        {filtered.length ? (
          <div className="table-scroll"><table className="management-table"><thead><tr><th>Etiqueta</th><th>Data de temporalidade</th><th>Documentos</th><th>Exportado em</th><th><span className="sr-only">Ações</span></th></tr></thead><tbody>{filtered.map((item) => <tr key={item.id}><td><strong>{item.etiqueta}</strong></td><td>{formatDate(item.data)}</td><td>{item.recordCount}</td><td>{new Intl.DateTimeFormat("pt-BR", { dateStyle: "short", timeStyle: "short" }).format(new Date(item.exportedAt))}</td><td><button className="delete-button" aria-label={`Remover ${item.etiqueta} do histórico`} onClick={() => onRemove(item.id)}><Trash2 size={16} /></button></td></tr>)}</tbody></table></div>
        ) : <EmptyState icon={<Boxes size={30} />} title={query ? "Nenhuma etiqueta encontrada" : "Nenhuma caixa exportada"} description={query ? "Ajuste a busca para localizar outra caixa." : "As caixas aparecerão aqui depois da primeira exportação em Excel."} />}
      </section>
    </div>
  );
}

function ReportsPage({ history: allHistory, onExport, isExporting }: { history: CompletedBox[]; onExport: (selected: CompletedBox[]) => void; isExporting: boolean }) {
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const invalidRange = !!from && !!to && from > to;
  const history = useMemo(() => invalidRange ? [] : allHistory.filter((item) => (!from || item.data >= from) && (!to || item.data <= to)), [allHistory, from, to, invalidRange]);
  const totalDocuments = history.reduce((sum, item) => sum + item.recordCount, 0);
  const average = history.length ? Math.round(totalDocuments / history.length) : 0;
  const largest = history.reduce<CompletedBox | null>((current, item) => !current || item.recordCount > current.recordCount ? item : current, null);
  const chartItems = history.slice(0, 6).reverse();
  const chartMax = Math.max(...chartItems.map((item) => item.recordCount), 1);

  return (
    <div className="page-enter operations-page">
      <PageHeader eyebrow="INDICADORES DA OPERAÇÃO" title="Relatórios" description="Analise somente os resultados gerados pelas caixas exportadas neste navegador." action={<button className="primary-button" disabled={!history.length || isExporting} onClick={() => onExport(history)}><Download size={18} /> Exportar relatório</button>} />
      <div className="report-filters"><label>Data de temporalidade: de<input type="date" value={from} onInput={(event) => setFrom(event.currentTarget.value)} onChange={(event) => setFrom(event.target.value)} /></label><label>Até<input type="date" value={to} onInput={(event) => setTo(event.currentTarget.value)} onChange={(event) => setTo(event.target.value)} /></label><button className="text-button" onClick={() => { setFrom(""); setTo(""); }}>Limpar período</button><span role="status">{invalidRange ? "A data inicial deve ser anterior à final." : history.length + (history.length === 1 ? " caixa no período" : " caixas no período")}</span></div>
      <section className="stats-grid">
        <StatCard icon={<Boxes size={20} />} tone="navy" value={history.length.toLocaleString("pt-BR")} label="Caixas processadas" note="Caixas no período selecionado" />
        <StatCard icon={<FileText size={20} />} tone="teal" value={totalDocuments.toLocaleString("pt-BR")} label="Documentos exportados" note="Documentos no período selecionado" />
        <StatCard icon={<BarChart3 size={20} />} tone="amber" value={average.toLocaleString("pt-BR")} label="Média por caixa" note="Documentos por exportação" />
        <StatCard icon={<Table2 size={20} />} tone="violet" value={largest?.recordCount.toLocaleString("pt-BR") ?? "0"} label="Maior caixa" note={largest?.etiqueta ?? "Sem dados disponíveis"} />
      </section>
      <section className="report-grid">
        <article className="report-card chart-card">
          <div className="panel-toolbar"><div><span className="eyebrow">VOLUME POR CAIXA</span><h2>Últimas exportações</h2></div><BarChart3 size={22} /></div>
          {chartItems.length ? <div className="bar-chart" aria-label="Quantidade de documentos por caixa">{chartItems.map((item) => <div className="bar-item" key={item.id}><div className="bar-track"><span style={{ height: `${Math.max(8, (item.recordCount / chartMax) * 100)}%` }}><b>{item.recordCount}</b></span></div><small>{item.etiqueta}</small></div>)}</div> : <EmptyState icon={<BarChart3 size={30} />} title="Relatório aguardando dados" description="Exporte uma caixa para iniciar os indicadores." />}
        </article>
        <article className="report-card privacy-report">
          <div className="privacy-icon"><ShieldCheck size={26} /></div>
          <span className="eyebrow">PRIVACIDADE DO RELATÓRIO</span><h2>Somente dados operacionais</h2>
          <p>Os indicadores utilizam etiqueta, data e quantidade de documentos. Nomes e matrículas não entram no histórico.</p>
          <ul><li><Check size={14} /> Sem envio ao servidor</li><li><Check size={14} /> Sem dados pessoais</li><li><Check size={14} /> Exclusão disponível nas configurações</li></ul>
        </article>
      </section>
    </div>
  );
}

function EmployeesPage({ employees, baseInfo, isImporting, onImport, onReset }: { employees: Employee[]; baseInfo: BaseInfo | null; isImporting: boolean; onImport: () => void; onReset: () => void }) {
  const [query, setQuery] = useState("");
  const normalizedQuery = normalizeMatricula(query);
  const results = useMemo(() => {
    if (!query.trim()) return [];
    const text = normalizeHeader(query);
    if (!text) return [];
    const matches: Employee[] = [];
    for (const employee of employees) {
      if ((normalizedQuery && employee.matricula.includes(normalizedQuery)) || normalizeHeader(employee.nome).includes(text)) matches.push(employee);
      if (matches.length === 25) break;
    }
    return matches;
  }, [employees, normalizedQuery, query]);

  return (
    <div className="page-enter operations-page">
      <PageHeader eyebrow="CONSULTA LOCAL" title="Base de funcionários" description="Gerencie a exportação do SOC e valide matrículas antes de iniciar uma caixa." action={<button className="primary-button" onClick={onImport} disabled={isImporting}>{isImporting ? <RefreshCw className="spin" size={18} /> : <Upload size={18} />}{baseInfo ? "Substituir base" : "Importar base"}</button>} />
      <section className={`base-management-card ${baseInfo ? "loaded" : ""}`}>
        <div className="base-management-status"><div className="soc-base-icon">{baseInfo ? <CheckCircle2 size={25} /> : <Database size={25} />}</div><div><span className="eyebrow">STATUS DA BASE</span><h2>{baseInfo ? "Base pronta para consulta" : "Nenhuma base carregada"}</h2><p>{baseInfo ? `${baseInfo.fileName} • aba ${baseInfo.sheetName}` : "Importe um arquivo .xlsx exportado do SOC."}</p></div></div>
        <div className="base-management-metrics"><div><span>Funcionários</span><strong>{baseInfo?.count.toLocaleString("pt-BR") ?? "0"}</strong></div><div><span>Importada em</span><strong>{baseInfo?.importedAt ?? "—"}</strong></div><div><span>Linhas idênticas repetidas</span><strong>{baseInfo?.duplicates.toLocaleString("pt-BR") ?? "0"}</strong></div></div>
        {baseInfo && <button className="base-reset-button" onClick={onReset}>Remover base</button>}
      </section>
      {baseInfo && <ImportSummary baseInfo={baseInfo} />}
      <section className="employee-search-card">
        <div className="panel-toolbar"><div><span className="eyebrow">PESQUISA</span><h2>Localizar funcionário</h2></div><span className="local-privacy"><Lock size={14} /> Consulta somente neste dispositivo</span></div>
        <div className="employee-search-box"><UserSearch size={21} /><input disabled={!baseInfo} value={query} onChange={(event) => setQuery(event.target.value)} aria-label="Pesquisar matrícula ou nome" placeholder={baseInfo ? "Digite matrícula ou nome" : "Importe a base do SOC para pesquisar"} /></div>
        {!query.trim() ? <EmptyState icon={<UserSearch size={30} />} title={baseInfo ? "Digite para pesquisar" : "Base necessária"} description={baseInfo ? "Os resultados aparecem conforme a matrícula ou o nome informado." : "A pesquisa será habilitada depois da importação."} /> : results.length ? <div className="employee-results">{results.map((employee) => <div className="employee-result" key={employee.matricula}><span>{employee.nome.split(" ").map((part) => part[0]).slice(0, 2).join("")}</span><div><strong>{employee.nome}</strong><small>Matrícula {employee.matricula}</small></div><CheckCircle2 size={18} /></div>)}</div> : <EmptyState icon={<Search size={30} />} title="Nenhum funcionário encontrado" description="Confira a matrícula ou o nome digitado." />}
        {results.length === 25 && <p className="results-note">Exibindo os primeiros 25 resultados. Refine a matrícula ou o nome para localizar outra pessoa.</p>}
      </section>
    </div>
  );
}

function ImportSummary({ baseInfo }: { baseInfo: BaseInfo }) {
  return <div className="import-summary">
    <span><CheckCircle2 size={16} /> {baseInfo.count.toLocaleString("pt-BR")} matrículas disponíveis na aba {baseInfo.sheetName}.</span>
    {!!baseInfo.duplicates && <span>{baseInfo.duplicates.toLocaleString("pt-BR")} linhas com a mesma matrícula e nome foram consolidadas.</span>}
    {!!baseInfo.conflicts && <strong><CircleAlert size={16} /> {baseInfo.conflicts.toLocaleString("pt-BR")} matrículas com nomes diferentes foram bloqueadas. Confira a base de origem.</strong>}
    {!!baseInfo.invalidRows && <span>{baseInfo.invalidRows.toLocaleString("pt-BR")} linhas sem matrícula válida ou nome não foram importadas.</span>}
    {baseInfo.sheetsFound > 1 && <span>Encontramos {baseInfo.sheetsFound} abas compatíveis. Foi usada a aba com mais matrículas válidas: {baseInfo.sheetName}.</span>}
    <small>A base permanece nesta sessão. Para refletir mudanças no SOC, importe uma exportação atualizada.</small>
  </div>;
}

function SettingsPage({ theme, onToggleTheme, preventDuplicates, onPreventDuplicates, baseInfo, historyCount, onImport, onResetBase, onClearHistory }: { theme: Theme; onToggleTheme: () => void; preventDuplicates: boolean; onPreventDuplicates: (value: boolean) => void; baseInfo: BaseInfo | null; historyCount: number; onImport: () => void; onResetBase: () => void; onClearHistory: () => void }) {
  return (
    <div className="page-enter operations-page">
      <PageHeader eyebrow="PREFERÊNCIAS LOCAIS" title="Configurações" description="Ajuste a experiência do Arquivo Ágil neste navegador." />
      <section className="settings-grid">
        <article className="settings-card"><div className="settings-card-head"><div className="settings-icon"><Sun size={21} /></div><div><h2>Aparência</h2><p>Alterne entre os modos claro e escuro.</p></div></div><button className="setting-row" onClick={onToggleTheme}><div><strong>Tema da interface</strong><span>{theme === "light" ? "Modo claro ativo" : "Modo escuro ativo"}</span></div><span className={`switch ${theme === "dark" ? "on" : ""}`}><i /></span></button></article>
        <article className="settings-card"><div className="settings-card-head"><div className="settings-icon"><ListFilter size={21} /></div><div><h2>Planilhamento</h2><p>Regras aplicadas durante a conferência.</p></div></div><button className="setting-row" onClick={() => onPreventDuplicates(!preventDuplicates)}><div><strong>Bloquear matrícula duplicada</strong><span>Evita incluir o mesmo funcionário duas vezes na caixa.</span></div><span className={`switch ${preventDuplicates ? "on" : ""}`}><i /></span></button><div className="setting-row static"><div><strong>Setor padrão</strong><span>Aplicado automaticamente em todas as linhas.</span></div><b>Saúde Ocupacional</b></div></article>
        <article className="settings-card"><div className="settings-card-head"><div className="settings-icon"><Database size={21} /></div><div><h2>Base do SOC</h2><p>A base permanece apenas na memória da página.</p></div></div><div className="setting-summary"><span>Status</span><strong>{baseInfo ? `${baseInfo.count.toLocaleString("pt-BR")} funcionários carregados` : "Nenhuma base carregada"}</strong></div><div className="settings-actions"><button className="base-import-button" onClick={onImport}><Upload size={16} />{baseInfo ? "Substituir" : "Importar"}</button>{baseInfo && <button className="danger-outline" onClick={onResetBase}>Remover base</button>}</div></article>
        <article className="settings-card"><div className="settings-card-head"><div className="settings-icon"><HardDrive size={21} /></div><div><h2>Dados locais</h2><p>Somente resumos operacionais são mantidos.</p></div></div><div className="setting-summary"><span>Histórico de caixas</span><strong>{historyCount} {historyCount === 1 ? "registro" : "registros"}</strong></div><p className="settings-note"><Lock size={14} /> Nomes e matrículas nunca são armazenados no histórico.</p><button className="danger-outline full" disabled={!historyCount} onClick={onClearHistory}><Trash2 size={16} /> Limpar histórico local</button></article>
      </section>
    </div>
  );
}

function EmptyState({ icon, title, description }: { icon: React.ReactNode; title: string; description: string }) {
  return <div className="operation-empty">{icon}<strong>{title}</strong><span>{description}</span></div>;
}

function StatCard({ icon, tone, value, label, note }: { icon: React.ReactNode; tone: string; value: string; label: string; note: string }) {
  return (
    <article className="stat-card">
      <div className={`stat-icon ${tone}`}>{icon}</div>
      <div className="stat-value">{value}</div>
      <div className="stat-label">{label}</div>
      <div className="stat-note"><span>↗</span>{note}</div>
    </article>
  );
}

type BatchProps = {
  step: 1 | 2 | 3;
  setStep: (step: 1 | 2 | 3) => void;
  label: string;
  setLabel: (label: string) => void;
  date: string;
  setDate: (date: string) => void;
  records: RecordRow[];
  scanStatus: ScanStatus;
  scanCopy: Record<ScanStatus, { title: string; detail: string }>;
  lastEmployee: Employee | null;
  manualInput: string;
  setManualInput: (value: string) => void;
  onManual: () => void;
  onRemove: (id: number) => void;
  onBack: () => void;
  onExport: () => void;
  isExporting: boolean;
  onReset: () => void;
  baseInfo: BaseInfo | null;
};

function BatchFlow(props: BatchProps) {
  const {
    step, setStep, label, setLabel, date, setDate, records, scanStatus, scanCopy,
    lastEmployee, manualInput, setManualInput, onManual,
    onRemove, onBack, onExport, onReset, baseInfo, isExporting,
  } = props;
  const busy = isExporting;

  return (
    <div className="page-enter batch-page">
      <button className="back-link" disabled={busy} onClick={step === 1 ? onBack : () => setStep((step - 1) as 1 | 2)}><ArrowLeft size={17} /> {step === 1 ? "Voltar à visão geral" : "Voltar à etapa anterior"}</button>

      <section className="batch-header">
        <div><span className="eyebrow">{baseInfo?.demo ? "DEMONSTRAÇÃO GUIADA" : "PLANILHAMENTO"}</span><h1>{step === 1 ? "Configure a caixa" : step === 2 ? `Consultando ${label}` : "Revise e exporte"}</h1><p>{step === 1 ? "Informe os dados gerais uma única vez antes de começar." : step === 2 ? "Cada matrícula localizada na base importada é adicionada à planilha." : "Confira os registros antes de gerar o arquivo final."}</p></div>
        <div className="stepper" aria-label={`Etapa ${step} de 3`}>
          {[1, 2, 3].map((item) => <div key={item} className={`step ${step === item ? "current" : ""} ${step > item ? "complete" : ""}`}><span>{step > item ? <Check size={14} /> : item}</span><small>{item === 1 ? "Caixa" : item === 2 ? "Leitura" : "Exportação"}</small></div>)}
        </div>
      </section>

      {step === 1 && (
        <section className="setup-layout">
          <article className="form-card">
            <div className="section-icon"><Boxes size={21} /></div>
            <h2>Dados da caixa</h2><p>Estes campos serão aplicados automaticamente a todos os documentos.</p>
            <label>Número da etiqueta <span>Obrigatório</span><div className="input-wrap"><FileText size={18} /><input value={label} maxLength={80} onChange={(e) => setLabel(e.target.value.toUpperCase())} placeholder="Digite a etiqueta real da caixa" /></div></label>
            <div className="field-row">
              <label>Data de temporalidade<div className="input-wrap"><Clock3 size={18} /><input type="date" value={date} onInput={(event) => setDate(event.currentTarget.value)} onChange={(e) => setDate(e.target.value)} /></div></label>
              <label>Setor<div className="input-wrap locked"><ShieldCheck size={18} /><input value="Saúde Ocupacional" readOnly /><span className="lock-label">FIXO</span></div></label>
            </div>
            <div className="form-tip"><ShieldCheck size={19} /><div><strong>Preenchimento inteligente</strong><span>Etiqueta, setor e data serão repetidos em cada linha da planilha.</span></div></div>
            <button className="primary-button form-submit" disabled={!label.trim() || !validDate(date) || !baseInfo || busy} onClick={() => setStep(2)}>Iniciar consultas <Search size={18} /></button>
          </article>

          <aside className="how-card">
            <span className="eyebrow">COMO FUNCIONA</span><h2>Do papel ao Excel<br />em três movimentos.</h2>
            <ol>
              <li><span><Keyboard size={18} /></span><div><strong>Digite a matrícula</strong><p>Use o número impresso no documento físico.</p></div></li>
              <li><span><Database size={18} /></span><div><strong>Consulta na base real</strong><p>O nome é localizado na planilha importada do SOC.</p></div></li>
              <li><span><FileSpreadsheet size={18} /></span><div><strong>Planilha pronta</strong><p>Revise os dados e exporte em .xlsx.</p></div></li>
            </ol>
            <div className="time-saved"><ShieldCheck size={20} /><div><strong>Teste controlado</strong><span>Comece com 3 a 5 documentos e confira todos os campos.</span></div></div>
          </aside>
        </section>
      )}

      {step === 2 && (
        <section className="scanner-layout">
          <div className="scanner-column">
            <article className={`scanner-card status-${scanStatus}`}>
              <div className="scanner-card-header"><div><span className="live-dot" /> {baseInfo?.demo ? "BASE FICTÍCIA DE DEMONSTRAÇÃO" : "BASE IMPORTADA CARREGADA"}</div><span>Consulta local</span></div>
              <div className="scanner-viewport">
                <div className="corner top-left" /><div className="corner top-right" /><div className="corner bottom-left" /><div className="corner bottom-right" />
                <div className="paper-preview">
                  <div className="paper-head"><div className="paper-symbol"><Users size={18} /></div><div><b>FICHA OCUPACIONAL</b><span>Documento de funcionário</span></div></div>
                  <div className="paper-lines"><i /><i /><i className="short" /></div>
                  <div className="ocr-target"><small>MATRÍCULA DO FUNCIONÁRIO</small><strong>{manualInput || lastEmployee?.matricula || "DIGITE ABAIXO"}</strong><span>Número informado</span></div>
                  <div className="paper-lines bottom"><i /><i className="short" /></div>
                </div>
                {busy && <div className="scan-beam" />}
                <div className="focus-label"><ShieldCheck size={15} /> CONSULTA NA BASE IMPORTADA</div>
              </div>
              <div className="scan-state" aria-live="polite">
                <div className={`state-icon ${scanStatus}`}>
                  {scanStatus === "idle" ? <Keyboard size={21} /> : scanStatus === "success" ? <Check size={22} /> : <ScanLine size={21} />}
                </div>
                <div><strong>{scanCopy[scanStatus].title}</strong><span>{scanCopy[scanStatus].detail}</span></div>
                {busy && scanStatus !== "success" && <div className="loading-dots"><i /><i /><i /></div>}
              </div>
              <button className="scan-button" disabled={busy || !manualInput.trim()} onClick={onManual}><Search size={19} />{busy ? "Consultando matrícula..." : "Consultar matrícula na base SOC"}</button>
              <p className="demo-hint">{baseInfo ? `Base ativa: ${baseInfo.fileName} • ${baseInfo.count.toLocaleString("pt-BR")} funcionários` : "Importe uma base do SOC para habilitar a consulta."}</p>
            </article>

            <article className="manual-card">
              <div><div className="manual-icon"><Keyboard size={18} /></div><div><strong>Matrícula do documento</strong><span>Digite exatamente como aparece no documento</span></div></div>
              <div className="manual-search"><Search size={17} /><input inputMode="text" autoFocus value={manualInput} aria-label="Matrícula para incluir na caixa" disabled={busy} maxLength={64} onChange={(e) => setManualInput(e.target.value.toUpperCase())} onKeyDown={(e) => e.key === "Enter" && onManual()} placeholder="Digite a matrícula" /><button disabled={busy || !manualInput.trim()} onClick={onManual}>Consultar</button></div>
            </article>
          </div>

          <aside className="batch-side">
            <article className="batch-summary">
              <div className="summary-head"><div><span className="eyebrow">CAIXA ATUAL</span><h2>{label}</h2></div><span className="open-pill"><i /> Em andamento</span></div>
              <dl><div><dt>Data de temporalidade</dt><dd>{formatDate(date)}</dd></div><div><dt>Setor</dt><dd>Saúde Ocupacional</dd></div></dl>
              <div className="big-count"><strong>{String(records.length).padStart(2, "0")}</strong><span>documentos<br />identificados</span></div>

              <p>Os registros ficam disponíveis apenas nesta sessão do navegador. A consulta usa a matrícula digitada; não há leitura por câmera.</p>
            </article>
            {lastEmployee && <article className="last-read"><div className="last-check"><Check size={17} /></div><div><span>ÚLTIMA LEITURA</span><strong>{lastEmployee.nome}</strong><small>Matrícula {lastEmployee.matricula}</small></div></article>}
            <button className="review-button" disabled={!records.length || busy} onClick={() => setStep(3)}>Revisar e finalizar caixa <ChevronRight size={18} /></button>
          </aside>

          <article className="records-card">
            <div className="records-heading"><div><h2>Registros desta caixa</h2><span>{records.length} {records.length === 1 ? "documento incluído" : "documentos incluídos"}</span></div><span className="autosave"><ShieldCheck size={15} /> Dados desta sessão</span></div>
            {records.length === 0 ? <div className="empty-records"><Search size={28} /><strong>Nenhum documento incluído ainda</strong><span>Digite uma matrícula real para consultar a base do SOC.</span></div> : <RecordsTable records={records} onRemove={onRemove} />}
          </article>
        </section>
      )}

      {step === 3 && (
        <section className="review-layout">
          <article className="review-card">
            <div className="review-toolbar"><div><div className="section-icon compact"><FileSpreadsheet size={20} /></div><div><h2>Prévia da planilha</h2><p>Confira etiqueta, nome, matrícula, setor e data em cada linha.</p></div></div><span className="ready-pill"><CheckCircle2 size={16} /> Pronta para exportar</span></div>
            <RecordsTable records={records} onRemove={onRemove} full />
            <div className="review-footer"><span><ShieldCheck size={16} /> {records.length} de {records.length} registros validados</span><span><strong>Tabela do Excel</strong> • filtros e linhas alternadas</span></div>
          </article>

          <aside className="export-card">
            <div className="export-icon"><FileSpreadsheet size={28} /></div><span className="eyebrow">RESUMO DA EXPORTAÇÃO</span><h2>{label}</h2>
            <div className="export-metrics"><div><span>Registros</span><strong>{records.length}</strong></div><div><span>Campos por linha</span><strong>5</strong></div><div><span>Pendências</span><strong className="zero">0</strong></div></div>
            <div className="excel-table-badge"><FileSpreadsheet size={15} /><div><strong>Tabela formatada</strong><span>Cabeçalho, filtros e linhas alternadas</span></div></div>
            <div className="export-fields"><strong>Campos incluídos</strong>{["Número da etiqueta", "Nome do funcionário", "Matrícula", "Setor", "Data de temporalidade"].map((field) => <span key={field}><Check size={13} />{field}</span>)}</div>
            <button className="primary-button export-button" disabled={!records.length || busy} onClick={onExport}><Download size={18} /> {isExporting ? "Gerando Excel..." : "Gerar planilha Excel"}</button>
            <button className="text-button" disabled={busy} onClick={onReset}><RotateCcw size={15} /> Reiniciar caixa</button>
          </aside>
        </section>
      )}
    </div>
  );
}

function RecordsTable({ records, onRemove, full = false }: { records: RecordRow[]; onRemove: (id: number) => void; full?: boolean }) {
  return (
    <div className="table-scroll">
      <table className={full ? "full-table" : ""}>
        <thead><tr><th>#</th>{full && <th>Etiqueta</th>}<th>Nome do funcionário</th><th>Matrícula</th>{full && <><th>Setor</th><th>Data de temporalidade</th></>}<th>Status</th><th><span className="sr-only">Ações</span></th></tr></thead>
        <tbody>{records.map((row, index) => <tr key={row.id}><td><span className="row-number">{String(index + 1).padStart(2, "0")}</span></td>{full && <td><strong className="cell-label">{row.etiqueta}</strong></td>}<td><div className="employee-cell"><span>{row.nome.split(" ").map((part) => part[0]).slice(0, 2).join("")}</span><strong>{row.nome}</strong></div></td><td><code>{row.matricula}</code></td>{full && <><td>{row.setor}</td><td>{formatDate(row.data)}</td></>}<td><span className="confirmed-pill"><Check size={12} />{row.status}</span></td><td><button className="delete-button" aria-label={`Remover ${row.nome}`} onClick={() => onRemove(row.id)}><Trash2 size={16} /></button></td></tr>)}</tbody>
      </table>
    </div>
  );
}
