import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { ChartJSNodeCanvas } from "chartjs-node-canvas";
import { Document, HeadingLevel, ImageRun, Packer, Paragraph } from "docx";
import JSZip from "jszip";

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

function resolveRelative(baseFile: string, target: string): string {
  const baseDir = path.posix.dirname(baseFile);
  return path.posix.normalize(path.posix.join(baseDir, target));
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .trim();
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

async function main() {
  const arquivo = requiredArg("--arquivo");
  const outArg = getArg("--out");

  const workbookBuffer = await fs.readFile(arquivo);
  const zip = await JSZip.loadAsync(workbookBuffer);

  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const wbRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!workbookXml || !wbRelsXml) {
    throw new Error("Estrutura XLSX invalida: workbook nao encontrado.");
  }

  const sheets = extractSheets(workbookXml);
  const wbRels = readRelationships(wbRelsXml);

  const chartCanvas = new ChartJSNodeCanvas({
    width: 960,
    height: 540,
    backgroundColour: "white",
  });

  const fileBase = path.basename(arquivo, path.extname(arquivo));
  const outFile = outArg ?? path.join(process.cwd(), "relatorios", `${fileBase}_graficos_completos.docx`);

  await fs.mkdir(path.dirname(outFile), { recursive: true });

  const children: Paragraph[] = [
    new Paragraph({ text: "Relatorio de Graficos do Workbook", heading: HeadingLevel.TITLE }),
    new Paragraph({ text: `Arquivo: ${path.basename(arquivo)}`, bullet: { level: 0 } }),
  ];

  let totalGraficos = 0;
  let totalPizza = 0;

  const palette = ["#457b9d", "#2a9d8f", "#e9c46a", "#f4a261", "#e76f51", "#8d99ae", "#6d597a"];

  for (const sheet of sheets) {
    const wbRel = wbRels.get(sheet.rId);
    if (!wbRel) continue;

    const sheetPath = resolveRelative("xl/workbook.xml", wbRel.target);
    const sheetRelsPath = path.posix.join(path.posix.dirname(sheetPath), "_rels", `${path.posix.basename(sheetPath)}.rels`);
    const sheetRelsXml = await zip.file(sheetRelsPath)?.async("string");
    if (!sheetRelsXml) continue;

    const sheetRels = readRelationships(sheetRelsXml);
    const drawingLinks = [...sheetRels.values()].filter((v) => v.type.endsWith("/drawing"));

    const chartsForSheet: Array<{ title: string; type: string; labels: string[]; values: number[] }> = [];

    for (const drawingLink of drawingLinks) {
      const drawingPath = resolveRelative(sheetPath, drawingLink.target);
      const drawingXml = await zip.file(drawingPath)?.async("string");
      if (!drawingXml) continue;

      const drawingRelPath = path.posix.join(
        path.posix.dirname(drawingPath),
        "_rels",
        `${path.posix.basename(drawingPath)}.rels`,
      );
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

        const type = extractChartType(chartXml);
        const title = extractChartTitle(chartXml);
        const { labels, values } = extractSeriesData(chartXml);
        if (!values.length) continue;

        chartsForSheet.push({ title, type, labels, values });
      }
    }

    if (!chartsForSheet.length) continue;

    children.push(new Paragraph({ text: `Aba ${sheet.name}`, heading: HeadingLevel.HEADING_1 }));

    for (let i = 0; i < chartsForSheet.length; i += 1) {
      const chart = chartsForSheet[i];
      totalGraficos += 1;
      if (["pieChart", "pie3DChart", "doughnutChart"].includes(chart.type)) totalPizza += 1;

      const colors = chart.labels.map((_, idx) => palette[idx % palette.length]);
      const image = await chartCanvas.renderToBuffer({
        type: chart.type === "doughnutChart" ? "doughnut" : "pie",
        data: {
          labels: chart.labels,
          datasets: [{ data: chart.values, backgroundColor: colors, borderColor: "#ffffff", borderWidth: 1 }],
        },
        options: {
          plugins: {
            legend: { position: "right" },
            title: { display: true, text: chart.title },
          },
        },
      });

      children.push(new Paragraph({ text: `${i + 1}. ${chart.title}`, heading: HeadingLevel.HEADING_2 }));
      children.push(
        new Paragraph({
          children: [new ImageRun({ data: image, transformation: { width: 620, height: 350 } })],
        }),
      );
    }
  }

  children.splice(
    2,
    0,
    new Paragraph({ text: `Total de graficos renderizados: ${totalGraficos}`, bullet: { level: 0 } }),
    new Paragraph({ text: `Total de graficos de pizza: ${totalPizza}`, bullet: { level: 0 } }),
    new Paragraph({ text: "" }),
  );

  const doc = new Document({ sections: [{ children }] });
  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(outFile, buffer);

  console.log(JSON.stringify({ arquivo, outFile, totalGraficos, totalPizza }, null, 2));
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro ao gerar DOCX completo de graficos:", message);
  process.exit(1);
});
