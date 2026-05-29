import { useState } from "react";
import { Plus, Trash2, Download, ChevronLeft, X } from "lucide-react";
import * as XLSX from "xlsx-js-style";
import { Button } from "@/components/ui/button";

type Row = {
  id: string;
  nome: string;
  numero: string;
  tamanho: string;
};

type ExportInfo = {
  equipe: string;
  projeto: string;
  data: string;
  cliente: string;
};

const ADULT_SIZES = ['PP', 'P', 'M', 'G', 'GG', 'XG', 'XXG'];
const KIDS_SIZES = ['INF 2', 'INF 4', 'INF 6', 'INF 8', 'INF 10', 'INF 12', 'INF 14'];
const BABY_SIZES = ['BABY PP', 'BABY P', 'BABY M', 'BABY G', 'BABY GG', 'BABY XG', 'BABY XXG'];
const ORDERED_SIZES = [...ADULT_SIZES, ...KIDS_SIZES, ...BABY_SIZES];

const createEmptyRow = (): Row => ({
  id: crypto.randomUUID(),
  nome: "",
  numero: "",
  tamanho: "",
});

const createInitialRows = () => Array.from({ length: 15 }, createEmptyRow);

const todayBR = () => {
  const d = new Date();
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
};

export default function App() {
  const [rows, setRows] = useState<Row[]>(createInitialRows);
  const [view, setView] = useState<'main' | 'summary'>('main');
  const [sizeModalRowId, setSizeModalRowId] = useState<string | null>(null);

  const [showExportModal, setShowExportModal] = useState(false);
  const [exportInfo, setExportInfo] = useState<ExportInfo>({
    equipe: "",
    projeto: "",
    data: todayBR(),
    cliente: "",
  });

  const updateRow = (id: string, field: keyof Row, value: string) => {
    setRows(rows.map(row => (row.id === id ? { ...row, [field]: value } : row)));
  };

  const addRow = () => setRows([...rows, createEmptyRow()]);

  const deleteRow = (id: string) => {
    setRows(rows.filter(row => row.id !== id));
  };

  const clearList = () => {
    if (window.confirm("Tem certeza que deseja limpar toda a lista?")) {
      setRows(createInitialRows());
    }
  };

  const openExportModal = () => {
    const validRows = rows.filter(r => r.nome.trim() || r.numero.trim() || r.tamanho.trim());
    if (validRows.length === 0) {
      alert("Preencha pelo menos uma linha antes de baixar a planilha.");
      return;
    }
    setShowExportModal(true);
  };

  const doExport = () => {
    const { equipe, projeto, data, cliente } = exportInfo;
    const validRows = rows.filter(r => r.nome.trim() || r.numero.trim() || r.tamanho.trim());

    const sDarkHeader = {
      fill: { fgColor: { rgb: "111111" } },
      font: { bold: true, color: { rgb: "FFFFFF" }, sz: 14, name: "Arial" },
      alignment: { horizontal: "center", vertical: "center" },
    };
    const sLabelCell = {
      fill: { fgColor: { rgb: "1C1C1C" } },
      font: { bold: true, color: { rgb: "FFFFFF" }, sz: 10, name: "Arial" },
      alignment: { horizontal: "left", vertical: "center" },
    };
    const sValueCell = {
      fill: { fgColor: { rgb: "F0F0F0" } },
      font: { bold: false, color: { rgb: "111111" }, sz: 10, name: "Arial" },
      alignment: { horizontal: "left", vertical: "center" },
    };
    const sTableHead = {
      fill: { fgColor: { rgb: "111111" } },
      font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11, name: "Arial" },
      alignment: { horizontal: "center", vertical: "center" },
    };
    const sDataName = {
      fill: { fgColor: { rgb: "F8F8F8" } },
      font: { color: { rgb: "111111" }, sz: 10, name: "Arial" },
      alignment: { horizontal: "left", vertical: "center" },
    };
    const sDataCenter = {
      fill: { fgColor: { rgb: "F8F8F8" } },
      font: { color: { rgb: "111111" }, sz: 10, name: "Arial" },
      alignment: { horizontal: "center", vertical: "center" },
    };
    const sTotalLabel = {
      fill: { fgColor: { rgb: "111111" } },
      font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11, name: "Arial" },
      alignment: { horizontal: "left", vertical: "center" },
    };
    const sTotalValue = {
      fill: { fgColor: { rgb: "18b35b" } },
      font: { bold: true, color: { rgb: "FFFFFF" }, sz: 13, name: "Arial" },
      alignment: { horizontal: "center", vertical: "center" },
    };
    const sEmpty = { fill: { fgColor: { rgb: "F0F0F0" } } };

    const ws1: Record<string, any> = {};
    let r = 0;

    const c1 = (row: number, col: number, value: any, style?: any, type?: string) => {
      const ref = XLSX.utils.encode_cell({ r: row, c: col });
      const t = type ?? (typeof value === "number" ? "n" : "s");
      ws1[ref] = { v: value, t, s: style };
    };

    c1(r, 0, "AM SPORTS - LISTA DE UNIFORMES", sDarkHeader);
    c1(r, 1, "", sDarkHeader);
    c1(r, 2, "", sDarkHeader);
    r++;
    r++;

    const infoRows: [string, string][] = [
      ["EQUIPE", equipe || "-"],
      ["PROJETO", projeto || "-"],
      ["DATA", data || "-"],
      ["CLIENTE", cliente || "-"],
    ];
    for (const [label, val] of infoRows) {
      c1(r, 0, label, sLabelCell);
      c1(r, 1, val, sValueCell);
      c1(r, 2, "", sEmpty);
      r++;
    }

    r++;

    const freezeAt = r + 1;
    c1(r, 0, "NOME", sTableHead);
    c1(r, 1, "NÚMERO", sTableHead);
    c1(r, 2, "TAMANHO", sTableHead);
    r++;

    for (const row of validRows) {
      c1(r, 0, row.nome, sDataName);
      c1(r, 1, row.numero, sDataCenter);
      c1(r, 2, row.tamanho, sDataCenter);
      r++;
    }

    r++;

    c1(r, 0, "TOTAL DE CAMISAS:", sTotalLabel);
    c1(r, 1, "", sTotalLabel);
    c1(r, 2, validRows.length, sTotalValue, "n");

    ws1["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r, c: 2 } });
    ws1["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
      { s: { r, c: 0 }, e: { r, c: 1 } },
    ];
    ws1["!cols"] = [{ wch: 34 }, { wch: 12 }, { wch: 14 }];
    ws1["!rows"] = Array.from({ length: r + 1 }, () => ({ hpt: 20 }));
    ws1["!rows"][0] = { hpt: 30 };
    ws1["!rows"][freezeAt - 1] = { hpt: 24 };
    ws1["!freeze"] = { xSplit: 0, ySplit: freezeAt };

    const ws2: Record<string, any> = {};
    let r2 = 0;

    const c2 = (row: number, col: number, value: any, style?: any, type?: string) => {
      const ref = XLSX.utils.encode_cell({ r: row, c: col });
      const t = type ?? (typeof value === "number" ? "n" : "s");
      ws2[ref] = { v: value, t, s: style };
    };

    c2(r2, 0, "RESUMO DE TAMANHOS", sDarkHeader);
    c2(r2, 1, "", sDarkHeader);
    r2++;
    r2++;

    c2(r2, 0, "TAMANHO", sTableHead);
    c2(r2, 1, "QUANTIDADE", sTableHead);
    r2++;

    const sizeCounts: Record<string, number> = {};
    for (const row of rows) {
      if (row.tamanho) sizeCounts[row.tamanho] = (sizeCounts[row.tamanho] || 0) + 1;
    }

    for (const size of ORDERED_SIZES) {
      if (sizeCounts[size]) {
        c2(r2, 0, size, sDataName);
        c2(r2, 1, sizeCounts[size], sDataCenter, "n");
        r2++;
      }
    }

    r2++;
    c2(r2, 0, "TOTAL GERAL:", sTotalLabel);
    c2(r2, 1, validRows.length, sTotalValue, "n");

    ws2["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: r2, c: 1 } });
    ws2["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    ws2["!cols"] = [{ wch: 20 }, { wch: 16 }];
    ws2["!rows"] = Array.from({ length: r2 + 1 }, () => ({ hpt: 20 }));
    ws2["!rows"][0] = { hpt: 30 };

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, "Lista Principal");
    XLSX.utils.book_append_sheet(wb, ws2, "Resumo de Tamanhos");

    const slug = equipe
      ? equipe
          .toLowerCase()
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
          .replace(/\s+/g, "-")
          .replace(/[^a-z0-9-]/g, "")
      : "uniformes";

    XLSX.writeFile(wb, `lista-${slug}.xlsx`);
    setShowExportModal(false);
  };

  const sizeCounts = rows.reduce((acc, row) => {
    if (row.tamanho) {
      acc[row.tamanho] = (acc[row.tamanho] || 0) + 1;
    }
    return acc;
  }, {} as Record<string, number>);

  const validSizesForSummary = ORDERED_SIZES.filter(size => sizeCounts[size] > 0);
  const totalShirts = Object.values(sizeCounts).reduce((sum, count) => sum + count, 0);

  return (
    <div className="min-h-[100dvh] w-full flex justify-center pb-12 bg-background">
      <div className="w-full max-w-[470px] flex flex-col gap-4 px-3 pt-6">

        <header className="text-center pb-2">
          <h1 className="font-[Teko] text-5xl font-bold tracking-tight text-white uppercase leading-none">Lista de Uniformes</h1>
          <p className="text-primary text-sm font-semibold tracking-widest uppercase">AM Sports</p>
        </header>

        {view === 'main' ? (
          <>
            <div className="grid grid-cols-2 gap-3 mb-2">
              <Button
                onClick={() => setView('summary')}
                variant="secondary"
                className="w-full font-bold h-12 uppercase tracking-wide bg-white/10 hover:bg-white/20 text-white"
              >
                Total de Camisa
              </Button>
              <Button
                onClick={openExportModal}
                className="w-full font-bold h-12 uppercase tracking-wide bg-primary hover:bg-primary/90 text-white"
              >
                <Download className="w-4 h-4 mr-2" />
                Baixar Planilha
              </Button>
            </div>

            <div className="bg-card rounded-xl shadow-xl overflow-hidden border border-white/5">
              <div className="grid grid-cols-[1fr_70px_98px_30px] gap-2 p-3 bg-muted/50 border-b border-border/40 text-xs font-bold text-muted-foreground uppercase tracking-wider">
                <div className="pl-1">Nome</div>
                <div className="text-center">Nº</div>
                <div className="text-center">Tam.</div>
                <div></div>
              </div>

              <div className="p-2 space-y-1.5 flex flex-col">
                {rows.map((row) => (
                  <div key={row.id} className="grid grid-cols-[1fr_70px_98px_30px] gap-2 items-center group">
                    <input
                      type="text"
                      value={row.nome}
                      onChange={(e) => updateRow(row.id, 'nome', e.target.value.toUpperCase())}
                      className="w-full h-10 px-3 bg-muted/30 border border-border/50 rounded-md text-sm font-medium uppercase text-card-foreground focus:outline-none focus:ring-2 focus:ring-primary/50 transition-all placeholder:text-muted-foreground/50"
                      placeholder="NOME"
                    />
                    <input
                      type="text"
                      inputMode="numeric"
                      value={row.numero}
                      onChange={(e) => updateRow(row.id, 'numero', e.target.value)}
                      className="w-full h-10 px-1 text-center bg-muted/30 border border-border/50 rounded-md text-sm font-bold text-card-foreground focus:outline-none focus:ring-2 focus:ring-primary/50 transition-all placeholder:text-muted-foreground/50"
                      placeholder="00"
                    />
                    <button
                      onClick={() => setSizeModalRowId(row.id)}
                      className="w-full h-10 px-1 text-center bg-muted/30 border border-border/50 rounded-md text-sm font-bold text-primary focus:outline-none focus:ring-2 focus:ring-primary/50 transition-all hover:bg-muted/50"
                    >
                      {row.tamanho || "SELECIONAR"}
                    </button>
                    <button
                      onClick={() => deleteRow(row.id)}
                      className="w-8 h-10 flex items-center justify-center text-muted-foreground hover:text-destructive transition-colors rounded-md hover:bg-destructive/10"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                ))}
              </div>

              <div className="p-3 bg-muted/20 border-t border-border/30 flex flex-col gap-3">
                <Button
                  onClick={addRow}
                  variant="outline"
                  className="w-full border-dashed border-2 hover:bg-muted/50"
                >
                  <Plus className="w-4 h-4 mr-2" />
                  Adicionar Linha
                </Button>
                <Button
                  onClick={clearList}
                  variant="ghost"
                  className="w-full text-destructive hover:text-destructive hover:bg-destructive/10"
                >
                  Limpar Lista
                </Button>
              </div>
            </div>

            <p className="text-center text-xs text-muted-foreground mt-4 font-medium">
              Clique no campo de tamanho para abrir as opções.
            </p>
          </>
        ) : (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-300">
            <Button
              onClick={() => setView('main')}
              variant="ghost"
              className="text-white hover:text-white hover:bg-white/10 mb-4 -ml-2"
            >
              <ChevronLeft className="w-4 h-4 mr-1" />
              Voltar para lista
            </Button>

            <div className="bg-card rounded-xl shadow-xl overflow-hidden border border-white/5 p-5">
              <h2 className="text-2xl font-bold text-card-foreground">Resumo de Camisas</h2>
              <p className="text-sm text-muted-foreground mb-6">Quantidade separada por tamanho</p>

              {validSizesForSummary.length > 0 ? (
                <div className="space-y-2 mb-6">
                  {validSizesForSummary.map(size => (
                    <div key={size} className="flex justify-between items-center p-3 rounded-lg bg-muted/30 border border-border/40">
                      <span className="font-bold text-card-foreground">{size}</span>
                      <span className="font-bold text-primary bg-primary/10 px-3 py-1 rounded-md">{sizeCounts[size]}</span>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="text-center py-8 text-muted-foreground font-medium border-2 border-dashed border-border/50 rounded-lg mb-6">
                  Nenhuma camisa na lista ainda.
                </div>
              )}

              <div className="pt-4 border-t border-border flex justify-between items-center">
                <span className="font-bold text-card-foreground uppercase tracking-wide text-sm">Total de camisas</span>
                <span className="text-3xl font-[Teko] font-bold text-primary leading-none">{totalShirts}</span>
              </div>
            </div>
          </div>
        )}

      </div>

      {showExportModal && (
        <div
          className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-end sm:items-center justify-center animate-in fade-in duration-200"
          onClick={() => setShowExportModal(false)}
        >
          <div
            className="bg-card w-full sm:max-w-[420px] sm:rounded-2xl rounded-t-2xl p-5 shadow-2xl animate-in slide-in-from-bottom-8 duration-300"
            onClick={e => e.stopPropagation()}
          >
            <div className="flex justify-between items-center mb-5">
              <h3 className="text-lg font-bold text-card-foreground uppercase tracking-wide">Informações da Planilha</h3>
              <button
                onClick={() => setShowExportModal(false)}
                className="p-2 -mr-2 text-muted-foreground hover:bg-muted rounded-full transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3 mb-6">
              {(
                [
                  { key: "equipe",  label: "Nome da Equipe",   placeholder: "Ex: Favela FC" },
                  { key: "projeto", label: "Nome do Projeto",  placeholder: "Ex: Temporada 2026" },
                  { key: "data",    label: "Data",             placeholder: "dd/mm/aaaa" },
                  { key: "cliente", label: "Nome do Cliente",  placeholder: "Ex: João Silva" },
                ] as { key: keyof ExportInfo; label: string; placeholder: string }[]
              ).map(({ key, label, placeholder }) => (
                <div key={key}>
                  <label className="block text-xs font-bold text-muted-foreground uppercase tracking-wider mb-1">
                    {label}
                  </label>
                  <input
                    type="text"
                    value={exportInfo[key]}
                    onChange={e => setExportInfo(prev => ({ ...prev, [key]: e.target.value }))}
                    placeholder={placeholder}
                    className="w-full h-10 px-3 bg-muted/30 border border-border/50 rounded-md text-sm font-medium text-card-foreground focus:outline-none focus:ring-2 focus:ring-primary/50 transition-all placeholder:text-muted-foreground/40"
                  />
                </div>
              ))}
            </div>

            <Button
              onClick={doExport}
              className="w-full h-12 font-bold uppercase tracking-wide bg-primary hover:bg-primary/90 text-white"
            >
              <Download className="w-4 h-4 mr-2" />
              Gerar e Baixar Planilha
            </Button>
          </div>
        </div>
      )}

      {sizeModalRowId && (
        <div
          className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-end sm:items-center justify-center animate-in fade-in duration-200"
          onClick={() => setSizeModalRowId(null)}
        >
          <div
            className="bg-card w-full sm:max-w-[400px] sm:rounded-2xl rounded-t-2xl p-5 shadow-2xl animate-in slide-in-from-bottom-8 duration-300 max-h-[85vh] overflow-y-auto"
            onClick={e => e.stopPropagation()}
          >
            <div className="flex justify-between items-center mb-6">
              <h3 className="text-lg font-bold text-card-foreground uppercase tracking-wide">Selecione o tamanho</h3>
              <button
                onClick={() => setSizeModalRowId(null)}
                className="p-2 -mr-2 text-muted-foreground hover:bg-muted rounded-full transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-6">
              <div>
                <h4 className="text-xs font-bold text-muted-foreground mb-3 uppercase tracking-wider">Adulto</h4>
                <div className="grid grid-cols-4 gap-2">
                  {ADULT_SIZES.map(size => (
                    <button
                      key={size}
                      onClick={() => {
                        updateRow(sizeModalRowId, 'tamanho', size);
                        setSizeModalRowId(null);
                      }}
                      className="py-2 px-1 text-sm font-bold rounded-md bg-muted/50 text-card-foreground border border-border/50 hover:bg-primary hover:text-white hover:border-primary transition-all"
                    >
                      {size}
                    </button>
                  ))}
                </div>
              </div>

              <div>
                <h4 className="text-xs font-bold text-muted-foreground mb-3 uppercase tracking-wider">Infantil</h4>
                <div className="grid grid-cols-4 gap-2">
                  {KIDS_SIZES.map(size => (
                    <button
                      key={size}
                      onClick={() => {
                        updateRow(sizeModalRowId, 'tamanho', size);
                        setSizeModalRowId(null);
                      }}
                      className="py-2 px-1 text-sm font-bold rounded-md bg-muted/50 text-card-foreground border border-border/50 hover:bg-primary hover:text-white hover:border-primary transition-all"
                    >
                      {size}
                    </button>
                  ))}
                </div>
              </div>

              <div>
                <h4 className="text-xs font-bold text-muted-foreground mb-3 uppercase tracking-wider">Baby Look</h4>
                <div className="grid grid-cols-4 gap-2">
                  {BABY_SIZES.map(size => (
                    <button
                      key={size}
                      onClick={() => {
                        updateRow(sizeModalRowId, 'tamanho', size);
                        setSizeModalRowId(null);
                      }}
                      className="py-2 px-1 text-[11px] font-bold rounded-md bg-muted/50 text-card-foreground border border-border/50 hover:bg-primary hover:text-white hover:border-primary transition-all leading-tight"
                    >
                      {size}
                    </button>
                  ))}
                </div>
              </div>
            </div>

            <div className="mt-8 pt-4 border-t border-border">
              <Button
                onClick={() => setSizeModalRowId(null)}
                variant="outline"
                className="w-full font-bold uppercase tracking-wide"
              >
                Fechar
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
