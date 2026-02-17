import "dotenv/config";
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { ConvexHttpClient } from "convex/browser";
import { Document, HeadingLevel, Packer, Paragraph } from "docx";
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

function buildCompetencia(ano: number, mes: number): string {
  return `${ano}-${String(mes).padStart(2, "0")}`;
}

function bullet(text: string): Paragraph {
  return new Paragraph({ text, bullet: { level: 0 } });
}

async function main() {
  const ano = requiredNumberArg("--ano");
  const mes = requiredNumberArg("--mes");
  const outputArg = getArg("--out");
  const regioesArg = getArg("--regioes");
  const regioes = (regioesArg ?? "1,2,3,11,12,13")
    .split(",")
    .map((r) => Number(r.trim()))
    .filter((r) => Number.isInteger(r));

  if (regioes.length === 0) {
    throw new Error("Nenhuma regiao valida informada em --regioes.");
  }

  const convexUrl = process.env.CONVEX_URL;
  if (!convexUrl) throw new Error("Defina CONVEX_URL no ambiente (.env ou shell).");

  const client = new ConvexHttpClient(convexUrl);
  const competencia = buildCompetencia(ano, mes);
  const outFile =
    outputArg ?? path.join(process.cwd(), "relatorios", `relatorio_consolidado_${competencia}.docx`);

  const resultados = [] as Array<{
    regiao: number;
    totalTrechos: number;
    totalKm: number;
    programadosNoMes: number;
    naoProgramadosNoMes: number;
    totalImportacoes: number;
    importacoesComErro: number;
    totalErros: number;
    porTipoFonte: Array<{ tipoFonte: string; totalTrechos: number; totalKm: number }>;
    topSrePorKm: Array<{ sre: string; km: number }>;
  }>;

  for (const regiao of regioes) {
    const [graficos, inconsistencias] = await Promise.all([
      client.query(api.trechos.obterGraficosCompetencia, { regiao, ano, mes }),
      client.query(api.trechos.obterInconsistenciasImportacao, { regiao, ano, mes }),
    ]);

    resultados.push({
      regiao,
      totalTrechos: graficos.kpis.totalTrechos,
      totalKm: graficos.kpis.totalKm,
      programadosNoMes: graficos.kpis.programadosNoMes,
      naoProgramadosNoMes: graficos.kpis.naoProgramadosNoMes,
      totalImportacoes: inconsistencias.resumo.totalImportacoes,
      importacoesComErro: inconsistencias.resumo.importacoesComErro,
      totalErros: inconsistencias.resumo.totalErros,
      porTipoFonte: graficos.series.porTipoFonte,
      topSrePorKm: graficos.series.topSrePorKm,
    });
  }

  const totais = resultados.reduce(
    (acc, item) => {
      acc.totalTrechos += item.totalTrechos;
      acc.totalKm += item.totalKm;
      acc.programadosNoMes += item.programadosNoMes;
      acc.naoProgramadosNoMes += item.naoProgramadosNoMes;
      acc.totalImportacoes += item.totalImportacoes;
      acc.importacoesComErro += item.importacoesComErro;
      acc.totalErros += item.totalErros;
      return acc;
    },
    {
      totalTrechos: 0,
      totalKm: 0,
      programadosNoMes: 0,
      naoProgramadosNoMes: 0,
      totalImportacoes: 0,
      importacoesComErro: 0,
      totalErros: 0,
    },
  );

  const children: Paragraph[] = [
    new Paragraph({ text: "Relatorio Tecnico Consolidado", heading: HeadingLevel.TITLE }),
    bullet(`Competencia: ${competencia}`),
    bullet(`Regioes: ${regioes.join(", ")}`),
    bullet(`Gerado em: ${new Date().toISOString()}`),
    new Paragraph({ text: "" }),

    new Paragraph({ text: "Totais Consolidados", heading: HeadingLevel.HEADING_1 }),
    bullet(`Total de trechos: ${totais.totalTrechos}`),
    bullet(`Total de extensao (km): ${totais.totalKm.toFixed(2)}`),
    bullet(`Programados no mes: ${totais.programadosNoMes}`),
    bullet(`Nao programados no mes: ${totais.naoProgramadosNoMes}`),
    bullet(`Total de importacoes: ${totais.totalImportacoes}`),
    bullet(`Importacoes com erro: ${totais.importacoesComErro}`),
    bullet(`Total de erros: ${totais.totalErros}`),
    new Paragraph({ text: "" }),
  ];

  for (const item of resultados) {
    children.push(new Paragraph({ text: `Regiao ${item.regiao}`, heading: HeadingLevel.HEADING_2 }));
    children.push(bullet(`Total de trechos: ${item.totalTrechos}`));
    children.push(bullet(`Total de extensao (km): ${item.totalKm.toFixed(2)}`));
    children.push(bullet(`Programados no mes: ${item.programadosNoMes}`));
    children.push(bullet(`Nao programados no mes: ${item.naoProgramadosNoMes}`));
    children.push(bullet(`Total de importacoes: ${item.totalImportacoes}`));
    children.push(bullet(`Importacoes com erro: ${item.importacoesComErro}`));
    children.push(bullet(`Total de erros: ${item.totalErros}`));
    children.push(new Paragraph({ text: "Distribuicao por tipo de fonte", heading: HeadingLevel.HEADING_3 }));
    for (const fonte of item.porTipoFonte) {
      children.push(bullet(`${fonte.tipoFonte}: ${fonte.totalTrechos} trechos / ${fonte.totalKm.toFixed(2)} km`));
    }
    children.push(new Paragraph({ text: "Top SRE por KM", heading: HeadingLevel.HEADING_3 }));
    for (const sre of item.topSrePorKm.slice(0, 5)) {
      children.push(bullet(`${sre.sre}: ${sre.km.toFixed(2)} km`));
    }
    children.push(bullet(`Relatorio detalhado: relatorio_regiao_${item.regiao}_${competencia}.docx`));
    children.push(new Paragraph({ text: "" }));
  }

  await fs.mkdir(path.dirname(outFile), { recursive: true });

  const doc = new Document({
    sections: [{ children }],
  });

  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(outFile, buffer);
  console.log(`Relatorio consolidado DOCX gerado em: ${outFile}`);
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error("Erro ao gerar relatorio consolidado DOCX:", message);
  process.exit(1);
});
