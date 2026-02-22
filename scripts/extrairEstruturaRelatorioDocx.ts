import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import JSZip from "jszip";

type BlockParagraph = {
  index: number;
  tipo: "paragrafo";
  texto: string;
  estilo: string;
  numerado: boolean;
  quebraPagina: boolean;
  headingNivel: number | null;
  imagens: Array<{ rId: string; target: string }>;
};

type BlockTable = {
  index: number;
  tipo: "tabela";
  linhas: string[][];
  totalLinhas: number;
  totalColunas: number;
};

type Block = BlockParagraph | BlockTable;

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

function decodeXmlText(text: string): string {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#10;/g, "\n")
    .replace(/&#13;/g, "\r");
}

function normalizeWhitespace(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

function readRelationships(xml: string): Map<string, string> {
  const map = new Map<string, string>();
  const rx = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/?>(?:<\/Relationship>)?/g;
  let m: RegExpExecArray | null;
  while ((m = rx.exec(xml)) !== null) {
    map.set(m[1], m[2]);
  }
  return map;
}

function findMatchingTagEnd(xml: string, startIndex: number, tagName: string): number {
  const rx = new RegExp(`</?${tagName}\\b[^>]*>`, "g");
  rx.lastIndex = startIndex;
  let depth = 0;
  let m: RegExpExecArray | null;

  while ((m = rx.exec(xml)) !== null) {
    const token = m[0];
    const closing = token.startsWith("</");
    const selfClosing = /\/>$/.test(token);

    if (!closing) depth += 1;
    if (closing) depth -= 1;
    if (!closing && selfClosing) depth -= 1;

    if (depth === 0) {
      return m.index + token.length;
    }
  }

  return -1;
}

function extractBodyXml(documentXml: string): string {
  const bodyOpen = documentXml.indexOf("<w:body");
  if (bodyOpen === -1) return "";

  const bodyOpenEnd = documentXml.indexOf(">", bodyOpen);
  if (bodyOpenEnd === -1) return "";

  const bodyClose = documentXml.indexOf("</w:body>", bodyOpenEnd + 1);
  if (bodyClose === -1) return "";

  return documentXml.slice(bodyOpenEnd + 1, bodyClose);
}

function extractTopLevelBlocks(bodyXml: string): Array<{ tipo: "paragrafo" | "tabela"; xml: string }> {
  const blocks: Array<{ tipo: "paragrafo" | "tabela"; xml: string }> = [];
  let cursor = 0;

  const findNext = (tag: "w:p" | "w:tbl", from: number): number => {
    const rx = new RegExp(`<${tag}(?:\\s|>)`, "g");
    rx.lastIndex = from;
    const m = rx.exec(bodyXml);
    return m ? m.index : -1;
  };

  while (cursor < bodyXml.length) {
    const pIdx = findNext("w:p", cursor);
    const tIdx = findNext("w:tbl", cursor);

    if (pIdx === -1 && tIdx === -1) break;
    const nextIdx =
      pIdx === -1 ? tIdx : tIdx === -1 ? pIdx : Math.min(pIdx, tIdx);

    if (nextIdx === tIdx) {
      const end = findMatchingTagEnd(bodyXml, tIdx, "w:tbl");
      if (end === -1) break;
      blocks.push({ tipo: "tabela", xml: bodyXml.slice(tIdx, end) });
      cursor = end;
      continue;
    }

    const pEnd = bodyXml.indexOf("</w:p>", pIdx);
    if (pEnd === -1) break;
    blocks.push({ tipo: "paragrafo", xml: bodyXml.slice(pIdx, pEnd + "</w:p>".length) });
    cursor = pEnd + "</w:p>".length;
  }

  return blocks;
}

function extractParagraphText(paragraphXml: string): string {
  const tokens: string[] = [];

  const brCount = [...paragraphXml.matchAll(/<w:br\b[^>]*\/>/g)].length;
  if (brCount > 0) {
    for (let i = 0; i < brCount; i += 1) tokens.push("\n");
  }

  const tabCount = [...paragraphXml.matchAll(/<w:tab\b[^>]*\/>/g)].length;
  if (tabCount > 0) {
    for (let i = 0; i < tabCount; i += 1) tokens.push("\t");
  }

  for (const m of paragraphXml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)) {
    tokens.push(decodeXmlText(m[1]));
  }

  return normalizeWhitespace(tokens.join(" "));
}

function extractParagraphStyle(paragraphXml: string): string {
  return paragraphXml.match(/<w:pStyle\b[^>]*\bw:val="([^"]+)"/)?.[1] ?? "Normal";
}

function extractHeadingLevel(style: string): number | null {
  const n = style.match(/heading\s*([1-9])/i)?.[1] ?? style.match(/titulo\s*([1-9])/i)?.[1];
  if (!n) return null;
  const parsed = Number(n);
  return Number.isInteger(parsed) ? parsed : null;
}

function inferHeadingFromText(text: string): { nivel: number; titulo: string } | null {
  const m = text.match(/^(\d+(?:\.\d+)*)\.\s+(.+)$/);
  if (!m) return null;
  const nivel = m[1].split(".").length;
  const titulo = m[2].replace(/\s+\d+$/, "").trim();
  return { nivel, titulo };
}

function extractParagraphImages(paragraphXml: string, rels: Map<string, string>): Array<{ rId: string; target: string }> {
  const ids = [...paragraphXml.matchAll(/<a:blip\b[^>]*\br:embed="([^"]+)"[^>]*\/?>(?:<\/a:blip>)?/g)].map((m) => m[1]);
  return ids.map((rId) => ({ rId, target: rels.get(rId) ?? "" }));
}

function extractTable(tableXml: string): string[][] {
  const rows: string[][] = [];
  const rowMatches = [...tableXml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)];

  for (const rowMatch of rowMatches) {
    const rowXml = rowMatch[0];
    const cells = [...rowXml.matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)].map((cellMatch) => {
      const cellXml = cellMatch[0];
      const paragraphs = [...cellXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map((p) => extractParagraphText(p[0])).filter(Boolean);
      return normalizeWhitespace(paragraphs.join(" | "));
    });
    rows.push(cells);
  }

  return rows;
}

async function main() {
  const docxPath = requiredArg("--docx");
  const outputArg = getArg("--out");

  const outPath =
    outputArg ??
    path.join(
      path.dirname(docxPath),
      `${path.basename(docxPath, path.extname(docxPath))}.estrutura.json`,
    );

  const bytes = await fs.readFile(docxPath);
  const zip = await JSZip.loadAsync(bytes);
  const documentXml = await zip.file("word/document.xml")?.async("string");
  const relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");

  if (!documentXml || !relsXml) {
    throw new Error("DOCX invalido: document.xml ou relationships ausentes.");
  }

  const rels = readRelationships(relsXml);
  const bodyXml = extractBodyXml(documentXml);
  const rawBlocks = extractTopLevelBlocks(bodyXml);

  const blocks: Block[] = [];
  for (const [idx, block] of rawBlocks.entries()) {
    if (block.tipo === "paragrafo") {
      const estilo = extractParagraphStyle(block.xml);
      const texto = extractParagraphText(block.xml);
      blocks.push({
        index: idx + 1,
        tipo: "paragrafo",
        texto,
        estilo,
        numerado: /<w:numPr\b/.test(block.xml),
        quebraPagina: /<w:br\b[^>]*\bw:type="page"[^>]*\/>/.test(block.xml) || /<w:lastRenderedPageBreak\b[^>]*\/>/.test(block.xml),
        headingNivel: extractHeadingLevel(estilo),
        imagens: extractParagraphImages(block.xml, rels),
      });
      continue;
    }

    const linhas = extractTable(block.xml);
    const totalColunas = linhas.reduce((max, row) => Math.max(max, row.length), 0);
    blocks.push({
      index: idx + 1,
      tipo: "tabela",
      linhas,
      totalLinhas: linhas.length,
      totalColunas,
    });
  }

  const paragrafos = blocks.filter((b) => b.tipo === "paragrafo") as BlockParagraph[];
  const tabelas = blocks.filter((b) => b.tipo === "tabela") as BlockTable[];
  const imagens = paragrafos.flatMap((p) => p.imagens);

  const mediaArquivos = zip
    .file(/^word\/media\//)
    .map((file) => file.name)
    .sort((a, b) => a.localeCompare(b, "pt-BR"));

  const payload = {
    arquivo: docxPath,
    geradoEm: new Date().toISOString(),
    estatisticas: {
      blocosTotal: blocks.length,
      paragrafos: paragrafos.length,
      tabelas: tabelas.length,
      imagensReferenciadas: imagens.length,
      arquivosMedia: mediaArquivos.length,
    },
    headings: paragrafos
      .flatMap((p) => {
        if (!p.texto) return [];
        if (p.headingNivel !== null) {
          return [{ index: p.index, nivel: p.headingNivel, texto: p.texto, origem: "estilo" as const }];
        }
        const inferred = inferHeadingFromText(p.texto);
        if (!inferred) return [];
        return [{ index: p.index, nivel: inferred.nivel, texto: inferred.titulo, origem: "numeracao" as const }];
      }),
    blocos: blocks,
  };

  await fs.mkdir(path.dirname(outPath), { recursive: true });
  await fs.writeFile(outPath, JSON.stringify(payload, null, 2), "utf-8");

  console.log(
    JSON.stringify(
      {
        ok: true,
        arquivoEntrada: docxPath,
        arquivoSaida: outPath,
        estatisticas: payload.estatisticas,
      },
      null,
      2,
    ),
  );
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro ao extrair estrutura do DOCX:", message);
  process.exit(1);
});
