import "dotenv/config";
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { ConvexHttpClient } from "convex/browser";
import { Document, HeadingLevel, Packer, Paragraph, PageBreak } from "docx";
import { api } from "../convex/_generated/api";

function getArg(flag: string): string | undefined {
  const idx = process.argv.indexOf(flag);
  if (idx === -1) return undefined;
  return process.argv[idx + 1];
}

function requiredNumberArg(flag: string): number {
  const value = getArg(flag);
  if (!value) throw new Error(`Parametro obrigatorio ausente: ${flag}`);
  const n = Number(value);
  if (!Number.isInteger(n)) throw new Error(`Parametro invalido em ${flag}: ${value}`);
  return n;
}

function competencia(ano: number, mes: number): string {
  return `${ano}-${String(mes).padStart(2, "0")}`;
}

function bullet(text: string): Paragraph {
  return new Paragraph({ text, bullet: { level: 0 } });
}

async function main() {
  const regiao = requiredNumberArg("--regiao");
  const ano = requiredNumberArg("--ano");
  const mes = requiredNumberArg("--mes");
  const outputArg = getArg("--out");

  const convexUrl = process.env.CONVEX_URL;
  if (!convexUrl) throw new Error("Defina CONVEX_URL no ambiente.");

  const client = new ConvexHttpClient(convexUrl);
  const [graficos, inconsistencias, resumoPav, resumoNaoPav] = await Promise.all([
    client.query(api.trechos.obterGraficosCompetencia, { regiao, ano, mes }),
    client.query(api.trechos.obterInconsistenciasImportacao, { regiao, ano, mes }),
    client.query(api.workbook.obterResumoWorkbook, { regiao, ano, mes, tipoFonte: "PAV" }),
    client.query(api.workbook.obterResumoWorkbook, { regiao, ano, mes, tipoFonte: "NAO_PAV" }),
  ]);

  const outFile =
    outputArg ?? path.join(process.cwd(), "relatorios", `relatorio_oficial_regiao_${regiao}_${competencia(ano, mes)}.docx`);
  await fs.mkdir(path.dirname(outFile), { recursive: true });

  const children: Paragraph[] = [
    new Paragraph({ text: `Relatorio Tecnico Oficial - Regiao ${regiao}`, heading: HeadingLevel.TITLE }),
    bullet(`Competencia: ${competencia(ano, mes)}`),
    bullet(`Data de emissao: ${new Date().toISOString()}`),
    new Paragraph({ text: "" }),

    new Paragraph({ text: "1. Resumo Executivo", heading: HeadingLevel.HEADING_1 }),
    bullet(`Total de trechos: ${graficos.kpis.totalTrechos}`),
    bullet(`Total de extensao (km): ${graficos.kpis.totalKm.toFixed(2)}`),
    bullet(`Programados no mes: ${graficos.kpis.programadosNoMes}`),
    bullet(`Nao programados no mes: ${graficos.kpis.naoProgramadosNoMes}`),
    new Paragraph({ text: "" }),

    new Paragraph({ text: "2. Controle de Importacoes", heading: HeadingLevel.HEADING_1 }),
    bullet(`Total de importacoes: ${inconsistencias.resumo.totalImportacoes}`),
    bullet(`Importacoes com erro: ${inconsistencias.resumo.importacoesComErro}`),
    bullet(`Total de erros: ${inconsistencias.resumo.totalErros}`),
    ...inconsistencias.porCodigo.slice(0, 10).map((i) => bullet(`${i.codigo}: ${i.total}`)),
    new Paragraph({ text: "" }),

    new Paragraph({ text: "3. Cobertura de Dados Complementares", heading: HeadingLevel.HEADING_1 }),
    bullet(`TT PAV (linhas): ${resumoPav.ttLinhas}`),
    bullet(`TT NAO_PAV (linhas): ${resumoNaoPav.ttLinhas}`),
    bullet(`Graficos PAV importados: ${resumoPav.graficos.total}`),
    bullet(`Graficos NAO_PAV importados: ${resumoNaoPav.graficos.total}`),
    new Paragraph({ text: "" }),

    new Paragraph({ text: "4. Top SRE por KM", heading: HeadingLevel.HEADING_1 }),
    ...graficos.series.topSrePorKm.slice(0, 10).map((s) => bullet(`${s.sre}: ${s.km.toFixed(2)} km`)),

    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ text: "5. Anexos e Evidencias", heading: HeadingLevel.HEADING_1 }),
    bullet(`Relatorio regional markdown: relatorio_regiao_${regiao}_${competencia(ano, mes)}.md`),
    bullet(`Relatorio regional DOCX: relatorio_regiao_${regiao}_${competencia(ano, mes)}.docx`),
    bullet(`Workbook completo PAV: pav_regiao${String(regiao).padStart(2, "0")}_workbook_completo.docx`),
    bullet(`Workbook completo NAO_PAV: nao_pav_regiao${String(regiao).padStart(2, "0")}_workbook_completo.docx`),
  ];

  const doc = new Document({ sections: [{ children }] });
  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(outFile, buffer);
  console.log(`Relatorio oficial DOCX gerado em: ${outFile}`);
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro ao gerar relatorio oficial DOCX:", message);
  process.exit(1);
});
