import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { ChartJSNodeCanvas } from "chartjs-node-canvas";
import JSZip from "jszip";
import * as XLSX from "xlsx";

type TipoFonte = "PAV" | "NAO_PAV";

type ChartData = {
  tipoFonte: TipoFonte;
  aba: string;
  titulo: string;
  tipoGrafico: string;
  trecho: string;
  labels: string[];
  valores: number[];
};

function getArg(flag: string): string | undefined {
  const idx = process.argv.indexOf(flag);
  if (idx === -1) return undefined;
  return process.argv[idx + 1];
}

function requiredArg(flag: string): string {
  const value = getArg(flag);
  if (!value) throw new Error(`Parametro obrigatorio ausente: ${flag}`);
  return value;
}

function normalizeText(text: string): string {
  return text
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function canonical(text: string): string {
  return normalizeText(text).replace(/\s+/g, "");
}

function decodeXmlText(text: string): string {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .trim();
}

function resolveRelative(baseFile: string, target: string): string {
  const baseDir = path.posix.dirname(baseFile);
  return path.posix.normalize(path.posix.join(baseDir, target));
}

function readRelationships(xml: string): Map<string, { type: string; target: string }> {
  const map = new Map<string, { type: string; target: string }>();
  const rx = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bType="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/?>(?:<\/Relationship>)?/g;
  let m: RegExpExecArray | null;
  while ((m = rx.exec(xml)) !== null) {
    map.set(m[1], { type: m[2], target: m[3] });
  }
  return map;
}

function extractSheets(workbookXml: string): Array<{ name: string; rId: string }> {
  const sheets: Array<{ name: string; rId: string }> = [];
  const rx = /<sheet\b[^>]*\bname="([^"]+)"[^>]*\br:id="([^"]+)"[^>]*\/?>(?:<\/sheet>)?/g;
  let m: RegExpExecArray | null;
  while ((m = rx.exec(workbookXml)) !== null) {
    sheets.push({ name: decodeXmlText(m[1]), rId: m[2] });
  }
  return sheets;
}

function extractChartRelationIds(drawingXml: string): string[] {
  const ids: string[] = [];
  const rx = /<c:chart\b[^>]*\br:id="([^"]+)"[^>]*\/?>(?:<\/c:chart>)?/g;
  let m: RegExpExecArray | null;
  while ((m = rx.exec(drawingXml)) !== null) {
    ids.push(m[1]);
  }
  return ids;
}

function extractChartTitle(chartXml: string): string {
  const matches = [...chartXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decodeXmlText(m[1]));
  return matches.join(" ").trim() || "Grafico";
}

function extractChartType(chartXml: string): string {
  const m = chartXml.match(/<c:(pieChart|pie3DChart|doughnutChart|barChart|lineChart|areaChart|radarChart|scatterChart)\b/);
  return m?.[1] ?? "unknown";
}

function extractCacheValues(cacheXml: string): string[] {
  const values: string[] = [];
  const rx = /<c:pt\b[^>]*>[\s\S]*?<c:v>([\s\S]*?)<\/c:v>[\s\S]*?<\/c:pt>/g;
  let m: RegExpExecArray | null;
  while ((m = rx.exec(cacheXml)) !== null) {
    values.push(decodeXmlText(m[1]));
  }
  return values;
}

function extractSeriesData(chartXml: string): { labels: string[]; values: number[] } {
  const ser = chartXml.match(/<c:ser\b[\s\S]*?<\/c:ser>/)?.[0] ?? chartXml;

  const strCache = ser.match(/<c:strCache\b[\s\S]*?<\/c:strCache>/)?.[0];
  const numCacheForCat = ser.match(/<c:cat\b[\s\S]*?<c:numCache\b[\s\S]*?<\/c:numCache>[\s\S]*?<\/c:cat>/)?.[0];
  const numCacheForVal = ser.match(/<c:val\b[\s\S]*?<c:numCache\b[\s\S]*?<\/c:numCache>[\s\S]*?<\/c:val>/)?.[0];

  const labels = strCache
    ? extractCacheValues(strCache)
    : numCacheForCat
      ? extractCacheValues(numCacheForCat)
      : [];

  const rawValues = numCacheForVal ? extractCacheValues(numCacheForVal) : [];
  const values = rawValues.map((v) => Number(v.replace(",", "."))).map((n) => (Number.isFinite(n) ? n : 0));

  const finalLabels = labels.length === values.length ? labels : values.map((_, i) => `Item ${i + 1}`);
  return { labels: finalLabels, values };
}

function valorCelula(sheet: XLSX.WorkSheet | undefined, endereco: string): string {
  if (!sheet) return "";
  const cell = sheet[endereco];
  return String(cell?.w ?? cell?.v ?? "").trim();
}

function resolverFormulaSimples(formula: string): { aba: string; celula: string } | null {
  const limpa = formula.trim();
  const m = limpa.match(/^'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)$/i);
  if (!m) return null;
  const aba = m[1].trim();
  const celula = `${m[2].toUpperCase()}${m[3]}`;
  return { aba, celula };
}

function extrairTrechoPorAba(workbook: XLSX.WorkBook): Record<string, string> {
  const trechoPorAba: Record<string, string> = {};
  for (const aba of workbook.SheetNames) {
    const sheet = workbook.Sheets[aba];
    if (!sheet) continue;
    const cell = sheet.B3 as XLSX.CellObject | undefined;
    let raw = valorCelula(sheet, "B3");

    if (!raw && cell?.f) {
      const ref = resolverFormulaSimples(cell.f);
      if (ref) raw = valorCelula(workbook.Sheets[ref.aba], ref.celula);
    }

    if (raw) trechoPorAba[aba] = raw;
  }
  return trechoPorAba;
}

async function lerChartsWorkbook(filePath: string, tipoFonte: TipoFonte): Promise<ChartData[]> {
  const bytes = await fs.readFile(filePath);
  const workbook = XLSX.read(bytes, { type: "buffer" });
  const trechoPorAba = extrairTrechoPorAba(workbook);

  const zip = await JSZip.loadAsync(bytes);
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const wbRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!workbookXml || !wbRelsXml) {
    throw new Error(`Estrutura XLSX invalida: ${filePath}`);
  }

  const sheets = extractSheets(workbookXml);
  const wbRels = readRelationships(wbRelsXml);

  const charts: ChartData[] = [];

  for (const sheet of sheets) {
    const wbRel = wbRels.get(sheet.rId);
    if (!wbRel) continue;

    const sheetPath = resolveRelative("xl/workbook.xml", wbRel.target);
    const sheetRelsPath = path.posix.join(path.posix.dirname(sheetPath), "_rels", `${path.posix.basename(sheetPath)}.rels`);
    const sheetRelsXml = await zip.file(sheetRelsPath)?.async("string");
    if (!sheetRelsXml) continue;

    const sheetRels = readRelationships(sheetRelsXml);
    const drawingLinks = [...sheetRels.values()].filter((v) => v.type.endsWith("/drawing"));

    let ordem = 0;
    for (const drawingLink of drawingLinks) {
      const drawingPath = resolveRelative(sheetPath, drawingLink.target);
      const drawingXml = await zip.file(drawingPath)?.async("string");
      if (!drawingXml) continue;

      const drawingRelPath = path.posix.join(path.posix.dirname(drawingPath), "_rels", `${path.posix.basename(drawingPath)}.rels`);
      const drawingRelsXml = await zip.file(drawingRelPath)?.async("string");
      if (!drawingRelsXml) continue;

      const drawingRels = readRelationships(drawingRelsXml);
      const chartRelIds = extractChartRelationIds(drawingXml);

      for (const chartRelId of chartRelIds) {
        const chartRel = drawingRels.get(chartRelId);
        if (!chartRel || !chartRel.type.endsWith("/chart")) continue;

        const chartPath = resolveRelative(drawingPath, chartRel.target);
        const chartXml = await zip.file(chartPath)?.async("string");
        if (!chartXml) continue;

        const { labels, values } = extractSeriesData(chartXml);
        if (!values.length) continue;

        ordem += 1;
        charts.push({
          tipoFonte,
          aba: sheet.name,
          titulo: extractChartTitle(chartXml),
          tipoGrafico: extractChartType(chartXml),
          trecho: trechoPorAba[sheet.name] ?? "",
          labels,
          valores: values,
        });
      }
    }
  }

  return charts.sort((a, b) => {
    const na = Number(a.aba);
    const nb = Number(b.aba);
    if (Number.isFinite(na) && Number.isFinite(nb) && na !== nb) return na - nb;
    if (a.aba !== b.aba) return a.aba.localeCompare(b.aba, "pt-BR");
    return a.titulo.localeCompare(b.titulo, "pt-BR");
  });
}

function extractParagraphText(paraXml: string): string {
  const texts = [...paraXml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((m) => decodeXmlText(m[1]));
  return texts.join(" ").replace(/\s+/g, " ").trim();
}

function isTrechoHeading(text: string): boolean {
  return /RODOVIA\s+[A-Z]{2}-?\d+/i.test(text);
}

function sanitizeTrechoHeading(text: string): string {
  const rodoviaStart = text.search(/RODOVIA\s+[A-Z]{2}-?\d+/i);
  const base = rodoviaStart >= 0 ? text.slice(rodoviaStart) : text;
  return base
    .replace(/PAGEREF\s+_Toc\d+\s+\\h\s*\d*$/i, "")
    .replace(/\s+/g, " ")
    .trim();
}

function detectTipoContext(text: string): TipoFonte | undefined {
  const n = normalizeText(text);
  if (n.includes("1 6 1") && n.includes("rodovias") && n.includes("pavimentadas")) return "PAV";
  if (n.includes("1 6 2") && n.includes("rodovias") && n.includes("nao pavimentadas")) return "NAO_PAV";
  return undefined;
}

function isGraficoCaption(text: string): boolean {
  const n = normalizeText(text);
  return n.includes("avaliacao do consorcio supervisor") && n.includes("condicoes de pista") && n.includes("extrapista");
}

function updateRelationshipTarget(relsXml: string, rid: string, newTarget: string): string {
  const rx = new RegExp(`(<Relationship\\b[^>]*\\bId=\"${rid}\"[^>]*\\bTarget=\")(?:[^\"]+)(\"[^>]*>)`);
  return relsXml.replace(rx, `$1${newTarget}$2`);
}

function pickChart(
  chartsByKey: Map<string, ChartData[]>,
  tipo: TipoFonte,
  trecho: string,
): { chart: ChartData | undefined; matchedKey: string | undefined } {
  const key = `${tipo}|${canonical(trecho)}`;
  const exact = chartsByKey.get(key);
  if (exact && exact.length > 0) {
    return { chart: exact.shift(), matchedKey: key };
  }

  const trechoCanon = canonical(trecho);
  for (const [candidateKey, list] of chartsByKey.entries()) {
    if (!candidateKey.startsWith(`${tipo}|`) || list.length === 0) continue;
    const candidateTrecho = candidateKey.slice(candidateKey.indexOf("|") + 1);
    if (candidateTrecho.includes(trechoCanon) || trechoCanon.includes(candidateTrecho)) {
      return { chart: list.shift(), matchedKey: candidateKey };
    }
  }

  return { chart: undefined, matchedKey: undefined };
}

function pickChartWithFallback(
  chartsByKey: Map<string, ChartData[]>,
  tipoPreferido: TipoFonte,
  trecho: string,
): { chart: ChartData | undefined; tipoUsado: TipoFonte | undefined } {
  const primary = pickChart(chartsByKey, tipoPreferido, trecho);
  if (primary.chart) return { chart: primary.chart, tipoUsado: tipoPreferido };

  const fallbackTipo: TipoFonte = tipoPreferido === "PAV" ? "NAO_PAV" : "PAV";
  const fallback = pickChart(chartsByKey, fallbackTipo, trecho);
  if (fallback.chart) return { chart: fallback.chart, tipoUsado: fallbackTipo };

  return { chart: undefined, tipoUsado: undefined };
}

async function renderChartImage(renderer: ChartJSNodeCanvas, chart: ChartData): Promise<Buffer> {
  const palette = ["#1f4e5f", "#2e7d6a", "#f2a65a", "#e76f51", "#6c757d", "#0f6e8c", "#84a98c", "#c44536"];
  const chartType = (() => {
    if (chart.tipoGrafico === "doughnutChart") return "doughnut" as const;
    if (chart.tipoGrafico === "barChart") return "bar" as const;
    if (chart.tipoGrafico === "lineChart" || chart.tipoGrafico === "areaChart") return "line" as const;
    return "pie" as const;
  })();

  return renderer.renderToBuffer({
    type: chartType,
    data: {
      labels: chart.labels,
      datasets: [
        {
          data: chart.valores,
          label: chart.titulo,
          backgroundColor: chart.labels.map((_, idx) => palette[idx % palette.length]),
          borderColor: chartType === "line" ? "#1f4e5f" : "#ffffff",
          borderWidth: chartType === "line" ? 2 : 1,
          fill: chart.tipoGrafico === "areaChart",
          tension: 0.2,
        },
      ],
    },
    options: {
      responsive: false,
      plugins: {
        legend: { position: chartType === "line" ? "top" : "right" },
      },
      scales:
        chartType === "bar" || chartType === "line"
          ? {
              y: {
                beginAtZero: true,
              },
            }
          : undefined,
    },
  });
}

async function main() {
  const templatePath = requiredArg("--template");
  const pavPath = requiredArg("--pav");
  const naoPavPath = requiredArg("--nao-pav");
  const outArg = getArg("--out");

  const outPath =
    outArg ??
    path.join(path.dirname(templatePath), `${path.basename(templatePath, path.extname(templatePath))} - graficos atualizados.docx`);

  const [pavCharts, naoPavCharts] = await Promise.all([
    lerChartsWorkbook(pavPath, "PAV"),
    lerChartsWorkbook(naoPavPath, "NAO_PAV"),
  ]);

  const chartsByKey = new Map<string, ChartData[]>();
  for (const c of [...pavCharts, ...naoPavCharts]) {
    if (!c.trecho) continue;
    const key = `${c.tipoFonte}|${canonical(c.trecho)}`;
    const list = chartsByKey.get(key) ?? [];
    list.push(c);
    chartsByKey.set(key, list);
  }

  const templateBuffer = await fs.readFile(templatePath);
  const zip = await JSZip.loadAsync(templateBuffer);
  const documentXml = await zip.file("word/document.xml")?.async("string");
  let relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");
  if (!documentXml || !relsXml) {
    throw new Error("Template DOCX invalido: document.xml ou relationships ausentes.");
  }

  const paragraphMatches = [...documentXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];
  const renderer = new ChartJSNodeCanvas({
    width: 1400,
    height: 860,
    backgroundColour: "white",
  });

  let tipoAtual: TipoFonte = "PAV";
  let trechoAtual = "";
  let aguardandoGrafico = false;
  let substituidos = 0;
  const faltantes: string[] = [];

  for (const paraMatch of paragraphMatches) {
    const paraXml = paraMatch[0];
    const text = extractParagraphText(paraXml);

    const tipoDetectado = text ? detectTipoContext(text) : undefined;
    if (tipoDetectado) tipoAtual = tipoDetectado;

    if (text && isTrechoHeading(text)) trechoAtual = sanitizeTrechoHeading(text);
    if (text && isGraficoCaption(text)) aguardandoGrafico = true;

    const embeds = [...paraXml.matchAll(/<a:blip[^>]*r:embed="([^"]+)"/g)].map((m) => m[1]);
    if (!embeds.length || !aguardandoGrafico) continue;

    for (const rid of embeds) {
      const { chart, tipoUsado } = pickChartWithFallback(chartsByKey, tipoAtual, trechoAtual);
      if (!chart) {
        faltantes.push(`${tipoAtual}: ${trechoAtual}`);
        aguardandoGrafico = false;
        continue;
      }

      if (tipoUsado) tipoAtual = tipoUsado;

      const image = await renderChartImage(renderer, chart);
      const mediaName = `grafico_atualizado_${String(substituidos + 1).padStart(3, "0")}.png`;
      const mediaPath = `word/media/${mediaName}`;
      zip.file(mediaPath, image);

      relsXml = updateRelationshipTarget(relsXml, rid, `media/${mediaName}`);
      substituidos += 1;
      aguardandoGrafico = false;
    }
  }

  zip.file("word/_rels/document.xml.rels", relsXml);
  const outputBuffer = await zip.generateAsync({ type: "nodebuffer" });
  await fs.mkdir(path.dirname(outPath), { recursive: true });
  await fs.writeFile(outPath, outputBuffer);

  const faltantesUnicos = Array.from(new Set(faltantes)).sort((a, b) => a.localeCompare(b, "pt-BR"));
  console.log(JSON.stringify({
    templatePath,
    outPath,
    chartsDisponiveis: pavCharts.length + naoPavCharts.length,
    chartsSubstituidos: substituidos,
    faltantes: faltantesUnicos,
  }, null, 2));
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro ao atualizar graficos do relatorio template:", message);
  process.exit(1);
});
