import { v } from "convex/values";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { api } from "./_generated/api";
import { action, mutation, query } from "./_generated/server";

type TipoFonte = "PAV" | "NAO_PAV";

const MAX_UPLOAD_BYTES = 25 * 1024 * 1024;

function normalize(text: string): string {
  return text
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function canonical(text: string): string {
  return normalize(text).replace(/\s+/g, "");
}

function buildCompetencia(ano: number, mes: number): string {
  return `${ano}-${String(mes).padStart(2, "0")}`;
}

function decodeBase64ToUint8Array(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function estimateBase64Bytes(base64: string): number {
  const padding = base64.endsWith("==") ? 2 : base64.endsWith("=") ? 1 : 0;
  return Math.floor((base64.length * 3) / 4) - padding;
}

function validateUploadSize(base64: string) {
  const bytes = estimateBase64Bytes(base64);
  if (bytes > MAX_UPLOAD_BYTES) {
    throw new Error(`Arquivo excede limite de ${Math.floor(MAX_UPLOAD_BYTES / (1024 * 1024))}MB.`);
  }
}

function toNumber(value: unknown): number {
  const raw = String(value ?? "").trim();
  if (!raw) return 0;
  const normalized = raw.includes(",") ? raw.replace(/\./g, "").replace(",", ".") : raw;
  const n = Number(normalized);
  return Number.isFinite(n) ? n : 0;
}

function getColumnIndex(headers: string[], aliases: string[]): number {
  const canon = headers.map((h) => canonical(h));
  const aliasCanon = aliases.map((a) => canonical(a));
  return canon.findIndex((h) => aliasCanon.includes(h));
}

async function parseCharts(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const wbXml = await zip.file("xl/workbook.xml")?.async("string");
  const wbRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!wbXml || !wbRelsXml) return [] as Array<any>;

  const relMap = new Map<string, string>();
  for (const m of wbRelsXml.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relMap.set(m[1], m[2]);
  }

  const sheets = [...wbXml.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)].map((m) => ({
    nome: m[1],
    rId: m[2],
  }));

  const resolve = (baseFile: string, target: string) => {
    const b = baseFile.split("/");
    b.pop();
    for (const part of target.split("/")) {
      if (!part || part === ".") continue;
      if (part === "..") b.pop();
      else b.push(part);
    }
    return b.join("/");
  };

  const charts: Array<{ aba: string; titulo: string; tipoGrafico: string; labels: string[]; valores: number[] }> = [];

  for (const s of sheets) {
    const sheetTarget = relMap.get(s.rId);
    if (!sheetTarget) continue;
    const sheetPath = resolve("xl/workbook.xml", sheetTarget);
    const sheetRelsPath = `${sheetPath.substring(0, sheetPath.lastIndexOf("/"))}/_rels/${sheetPath.substring(sheetPath.lastIndexOf("/") + 1)}.rels`;
    const sheetRelsXml = await zip.file(sheetRelsPath)?.async("string");
    if (!sheetRelsXml) continue;

    const drawTargets = [...sheetRelsXml.matchAll(/<Relationship\b[^>]*Type="[^"]*\/drawing"[^>]*Target="([^"]+)"/g)].map(
      (m) => m[1],
    );

    for (const drawTarget of drawTargets) {
      const drawingPath = resolve(sheetPath, drawTarget);
      const drawingXml = await zip.file(drawingPath)?.async("string");
      if (!drawingXml) continue;

      const drawingRelsPath = `${drawingPath.substring(0, drawingPath.lastIndexOf("/"))}/_rels/${drawingPath.substring(drawingPath.lastIndexOf("/") + 1)}.rels`;
      const drawingRelsXml = await zip.file(drawingRelsPath)?.async("string");
      if (!drawingRelsXml) continue;

      const drawingRelMap = new Map<string, string>();
      for (const m of drawingRelsXml.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
        drawingRelMap.set(m[1], m[2]);
      }

      const chartIds = [...drawingXml.matchAll(/<c:chart\b[^>]*r:id="([^"]+)"/g)].map((m) => m[1]);
      for (const chartId of chartIds) {
        const chartTarget = drawingRelMap.get(chartId);
        if (!chartTarget) continue;
        const chartPath = resolve(drawingPath, chartTarget);
        const chartXml = await zip.file(chartPath)?.async("string");
        if (!chartXml) continue;

        const title = [...chartXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1].trim()).join(" ") || "Grafico";
        const type = chartXml.match(/<c:(pieChart|pie3DChart|doughnutChart|barChart|lineChart|areaChart|radarChart)\b/)?.[1] ?? "unknown";

        const labels = [...chartXml.matchAll(/<c:strCache>[\s\S]*?<c:pt[^>]*>[\s\S]*?<c:v>([\s\S]*?)<\/c:v>/g)].map(
          (m) => m[1].trim(),
        );
        const values = [...chartXml.matchAll(/<c:val>[\s\S]*?<c:numCache>[\s\S]*?<c:pt[^>]*>[\s\S]*?<c:v>([\s\S]*?)<\/c:v>/g)].map(
          (m) => toNumber(m[1]),
        );
        if (!values.length) continue;

        charts.push({
          aba: s.nome,
          titulo: title,
          tipoGrafico: type,
          labels: labels.length === values.length ? labels : values.map((_, i) => `Item ${i + 1}`),
          valores: values,
        });
      }
    }
  }

  return charts;
}

export const importarWorkbookComplementar = action({
  args: {
    sessionToken: v.string(),
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),
    arquivoOrigem: v.string(),
    arquivoBase64: v.string(),
    limparAntes: v.optional(v.boolean()),
    operador: v.optional(v.string()),
    perfil: v.optional(v.string()),
  },
  handler: async (ctx, args): Promise<any> => {
    const sessao: any = await ctx.runQuery(api.auth.me, { sessionToken: args.sessionToken });
    if (!sessao) throw new Error("Sessao invalida ou expirada.");
    if (!["OPERADOR", "GESTOR", "ADMIN"].includes(sessao.usuario.perfil)) {
      throw new Error("Permissao insuficiente para esta operacao.");
    }
    validateUploadSize(args.arquivoBase64);
    const bytes = decodeBase64ToUint8Array(args.arquivoBase64);
    const workbook = XLSX.read(bytes, { type: "array" });
    const competencia = buildCompetencia(args.ano, args.mes);

    const ttRows: Array<{
      numeroTrecho?: string;
      trecho: string;
      grupo: string;
      classificacao: string;
      valor: number;
    }> = [];

    const ttSheet = workbook.Sheets.TT;
    if (ttSheet) {
      const rows = XLSX.utils.sheet_to_json<Array<unknown>>(ttSheet, { header: 1, raw: false, defval: "" });
      const headerIdx = rows.findIndex((r) => {
        const h = r.map((c) => canonical(String(c ?? "")));
        return h.includes("trecho") && h.includes("grupo") && (h.includes("classificacao") || h.includes("classificação")) && h.includes("valor");
      });

      if (headerIdx >= 0) {
        const header = rows[headerIdx].map((c) => String(c ?? ""));
        const idxNumero = getColumnIndex(header, ["Nº do Trecho", "Numero Trecho", "N do Trecho"]);
        const idxTrecho = getColumnIndex(header, ["Trecho"]);
        const idxGrupo = getColumnIndex(header, ["Grupo"]);
        const idxClass = getColumnIndex(header, ["Classificação", "Classificacao"]);
        const idxValor = getColumnIndex(header, ["Valor"]);

        if (idxTrecho >= 0 && idxGrupo >= 0 && idxClass >= 0 && idxValor >= 0) {
          const dataRows = rows.slice(headerIdx + 1).filter((r) => r.some((c) => String(c ?? "").trim() !== ""));
          for (const r of dataRows) {
            const trecho = String(r[idxTrecho] ?? "").trim();
            const grupo = String(r[idxGrupo] ?? "").trim();
            const classificacao = String(r[idxClass] ?? "").trim();
            const valor = toNumber(r[idxValor]);
            if (!trecho || !grupo || !classificacao) continue;
            ttRows.push({
              numeroTrecho: idxNumero >= 0 ? String(r[idxNumero] ?? "").trim() : undefined,
              trecho,
              grupo,
              classificacao,
              valor,
            });
          }
        }
      }
    }

    const charts = await parseCharts(bytes);

    return (ctx as any).runMutation("workbook:persistirWorkbookComplementar", {
      regiao: args.regiao,
      ano: args.ano,
      mes: args.mes,
      competencia,
      tipoFonte: args.tipoFonte,
      arquivoOrigem: args.arquivoOrigem,
      limparAntes: args.limparAntes,
      operador: args.operador ?? sessao.nome,
      perfil: args.perfil ?? sessao.perfil,
      ttRows,
      charts,
    });
  },
});

export const persistirWorkbookComplementar = mutation({
  args: {
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    competencia: v.string(),
    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),
    arquivoOrigem: v.string(),
    limparAntes: v.optional(v.boolean()),
    operador: v.optional(v.string()),
    perfil: v.optional(v.string()),
    ttRows: v.array(
      v.object({
        numeroTrecho: v.optional(v.string()),
        trecho: v.string(),
        grupo: v.string(),
        classificacao: v.string(),
        valor: v.number(),
      }),
    ),
    charts: v.array(
      v.object({
        aba: v.string(),
        titulo: v.string(),
        tipoGrafico: v.string(),
        labels: v.array(v.string()),
        valores: v.array(v.number()),
      }),
    ),
  },
  handler: async (ctx, args) => {
    const competencia = buildCompetencia(args.ano, args.mes);
    void competencia;

    if (args.limparAntes === true) {
      const antigosTt = await ctx.db
        .query("ttAvaliacoes")
        .withIndex("by_regiao_ano_mes_tipo", (q) =>
          q.eq("regiao", args.regiao).eq("ano", args.ano).eq("mes", args.mes).eq("tipoFonte", args.tipoFonte),
        )
        .collect();
      for (const a of antigosTt) await ctx.db.delete(a._id);

      const antigosGraficos = await ctx.db
        .query("workbookGraficos")
        .withIndex("by_regiao_ano_mes_tipo", (q) =>
          q.eq("regiao", args.regiao).eq("ano", args.ano).eq("mes", args.mes).eq("tipoFonte", args.tipoFonte),
        )
        .collect();
      for (const a of antigosGraficos) await ctx.db.delete(a._id);
    }

    let ttInseridos = 0;

    for (const row of args.ttRows) {
      await ctx.db.insert("ttAvaliacoes", {
        regiao: args.regiao,
        ano: args.ano,
        mes: args.mes,
        competencia: args.competencia,
        tipoFonte: args.tipoFonte,
        numeroTrecho: row.numeroTrecho,
        trecho: row.trecho,
        grupo: row.grupo,
        classificacao: row.classificacao,
        valor: row.valor,
        arquivoOrigem: args.arquivoOrigem,
        importadoEm: Date.now(),
      });
      ttInseridos += 1;
    }
    let graficosInseridos = 0;
    for (let i = 0; i < args.charts.length; i += 1) {
      const c = args.charts[i];
      await ctx.db.insert("workbookGraficos", {
        regiao: args.regiao,
        ano: args.ano,
        mes: args.mes,
        competencia: args.competencia,
        tipoFonte: args.tipoFonte,
        aba: c.aba,
        ordem: i + 1,
        titulo: c.titulo,
        tipoGrafico: c.tipoGrafico,
        labels: c.labels,
        valores: c.valores,
        arquivoOrigem: args.arquivoOrigem,
        importadoEm: Date.now(),
      });
      graficosInseridos += 1;
    }

    await ctx.db.insert("auditoriaEventos", {
      acao: "WORKBOOK_COMPLEMENTAR_IMPORTADO",
      entidade: "workbook",
      regiao: args.regiao,
      ano: args.ano,
      mes: args.mes,
      competencia,
      operador: args.operador,
      perfil: args.perfil,
      detalhes: `${args.arquivoOrigem}; TT=${ttInseridos}; Graficos=${graficosInseridos}`,
      criadoEm: Date.now(),
    });

    return {
      ok: true,
      competencia,
      ttInseridos,
      graficosInseridos,
    };
  },
});

export const obterResumoWorkbook = query({
  args: {
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),
  },
  handler: async (ctx, args) => {
    const competencia = buildCompetencia(args.ano, args.mes);
    const tt = await ctx.db
      .query("ttAvaliacoes")
      .withIndex("by_regiao_ano_mes_tipo", (q) =>
        q.eq("regiao", args.regiao).eq("ano", args.ano).eq("mes", args.mes).eq("tipoFonte", args.tipoFonte),
      )
      .collect();

    const graficos = await ctx.db
      .query("workbookGraficos")
      .withIndex("by_regiao_ano_mes_tipo", (q) =>
        q.eq("regiao", args.regiao).eq("ano", args.ano).eq("mes", args.mes).eq("tipoFonte", args.tipoFonte),
      )
      .collect();

    return {
      competencia,
      ttLinhas: tt.length,
      graficos: {
        total: graficos.length,
        porAba: Array.from(
          graficos.reduce((acc, g) => {
            acc.set(g.aba, (acc.get(g.aba) ?? 0) + 1);
            return acc;
          }, new Map<string, number>()),
        ).map(([aba, total]) => ({ aba, total })),
      },
    };
  },
});

export const listarGraficosWorkbook = query({
  args: {
    sessionToken: v.string(),
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),
  },
  handler: async (ctx, args) => {
    const sessao = await ctx.runQuery(api.auth.me, { sessionToken: args.sessionToken });
    if (!sessao) throw new Error("Sessao invalida ou expirada.");

    const graficos = await ctx.db
      .query("workbookGraficos")
      .withIndex("by_regiao_ano_mes_tipo", (q) =>
        q.eq("regiao", args.regiao).eq("ano", args.ano).eq("mes", args.mes).eq("tipoFonte", args.tipoFonte),
      )
      .collect();

    const payload = graficos.map((g) => {
      const total = g.valores.reduce((acc, v) => acc + v, 0);
      const series = g.labels.map((label, i) => {
        const valor = g.valores[i] ?? 0;
        const percentual = total > 0 ? Number(((valor / total) * 100).toFixed(1)) : 0;
        return { label, valor, percentual };
      });

      return {
        id: g._id,
        aba: g.aba,
        ordem: g.ordem,
        titulo: g.titulo,
        tipoGrafico: g.tipoGrafico,
        labels: g.labels,
        valores: g.valores,
        total,
        series,
      };
    });

    return payload.sort((a, b) => {
      if (a.aba === b.aba) return a.ordem - b.ordem;
      const na = Number(a.aba);
      const nb = Number(b.aba);
      if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
      if (Number.isFinite(na)) return -1;
      if (Number.isFinite(nb)) return 1;
      return a.aba.localeCompare(b.aba, "pt-BR");
    });
  },
});
