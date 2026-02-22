import { v } from "convex/values";
import * as XLSX from "xlsx";
import { internal } from "./_generated/api";
import { internalMutation, mutation, query } from "./_generated/server";
import { requireSession } from "./security";

type TipoFonte = "PAV" | "NAO_PAV";

const MAX_UPLOAD_BYTES = 25 * 1024 * 1024;

const mesesProgramacao = [
  { aliases: ["jul", "july", "jul-"], mes: 7 },
  { aliases: ["ago", "aug", "aug-"], mes: 8 },
  { aliases: ["set", "sep", "sept", "sep-"], mes: 9 },
  { aliases: ["out", "oct", "oct-"], mes: 10 },
  { aliases: ["nov", "nov-"], mes: 11 },
  { aliases: ["dez", "dec", "dec-"], mes: 12 },
] as const;

function toTrimmedString(value: unknown): string | undefined {
  if (value === null || value === undefined) return undefined;
  const s = String(value).trim();
  return s.length > 0 ? s : undefined;
}

function toNumber(value: unknown): number | undefined {
  if (value === null || value === undefined || value === "") return undefined;
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const raw = String(value).trim();
  const hasDot = raw.includes(".");
  const hasComma = raw.includes(",");

  let normalized = raw;
  if (hasDot && hasComma) {
    normalized = raw.replace(/\./g, "").replace(",", ".");
  } else if (hasComma) {
    normalized = raw.replace(",", ".");
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function toBoolean(value: unknown): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value > 0;

  const s = String(value ?? "").trim().toLowerCase();
  return ["x", "sim", "s", "true", "1", "ok"].includes(s);
}

function normalizeKey(key: string): string {
  return key
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function getFirst(row: Record<string, unknown>, keys: string[]): unknown {
  for (const key of keys) {
    if (key in row) return row[key];
  }

  const rowKeys = Object.keys(row);
  const normalizedRowKeys = new Map<string, string>();
  for (const key of rowKeys) normalizedRowKeys.set(normalizeKey(key), key);

  for (const key of keys) {
    const found = normalizedRowKeys.get(normalizeKey(key));
    if (found) return row[found];
  }

  return undefined;
}

function buildProgramacao(row: Record<string, unknown>, anoBase: number): Record<string, boolean> {
  const result: Record<string, boolean> = {};

  for (const [rawKey, rawValue] of Object.entries(row)) {
    const key = normalizeKey(rawKey);
    const monthInfo = mesesProgramacao.find((m) => m.aliases.some((a) => key.includes(a)));
    if (!monthInfo) continue;

    const mes = String(monthInfo.mes).padStart(2, "0");
    result[`${anoBase}-${mes}`] = toBoolean(rawValue);
  }

  return result;
}

function buildCompetencia(ano: number, mes: number): string {
  return `${ano}-${String(mes).padStart(2, "0")}`;
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

function canonicalHeader(text: string): string {
  return normalizeKey(text).replace(/\s+/g, "");
}

function mapHeaderKey(header: string, idx: number): string {
  const n = canonicalHeader(header);
  if (n === "lote") return "LOTE";
  if (n === "n" || n === "numero" || n === "numerotrecho") return "NUMERO";
  if (n === "regiaodeconservacao") return "REGIAO_CONSERVACAO";
  if (n === "cidadesede") return "CIDADE_SEDE";
  if (n === "trecho" || n === "trechos") return "TRECHO";
  if (n === "sre") return "SRE";
  if (n === "subtrechos" || n === "subtrecho" || n === "sbutrecho" || n === "segmentos") return "SUBTRECHOS";
  if (n === "extkm" || n === "extensao" || n === "extensaokm") return "EXT_KM";
  if (n === "tipo") return "TIPO";
  if (n.startsWith("jul")) return "JUL";
  if (n.startsWith("aug") || n.startsWith("ago")) return "AGO";
  if (n.startsWith("sep") || n.startsWith("set")) return "SET";
  if (n.startsWith("oct") || n.startsWith("out")) return "OUT";
  if (n.startsWith("nov")) return "NOV";
  if (n.startsWith("dec") || n.startsWith("dez")) return "DEZ";
  return `COL_${idx + 1}`;
}

function decodeBase64ToUint8Array(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function lerLinhasTrechosDeArquivoBase64(arquivoBase64: string): Record<string, unknown>[] {
  const bytes = decodeBase64ToUint8Array(arquivoBase64);
  const workbook = XLSX.read(bytes, { type: "array" });
  const sheet = workbook.Sheets.Trechos;
  if (!sheet) throw new Error('A aba "Trechos" nao existe no arquivo enviado.');

  const matriz = XLSX.utils.sheet_to_json<Array<unknown>>(sheet, {
    header: 1,
    raw: false,
    defval: "",
  });

  const headerIdx = matriz.findIndex((row) => {
    const tokens = row.map((cell) => canonicalHeader(String(cell ?? "")));
    return tokens.some((cell) => cell === "trecho" || cell === "trechos") && tokens.includes("sre");
  });

  if (headerIdx === -1) throw new Error("Cabecalho da aba Trechos nao encontrado.");

  const header = matriz[headerIdx].map((c) => String(c ?? "").trim());
  const linhas = matriz
    .slice(headerIdx + 1)
    .filter((row) => row.some((c) => String(c ?? "").trim() !== ""))
    .map((row) => {
      const obj: Record<string, unknown> = {};
      for (let i = 0; i < header.length; i += 1) {
        obj[mapHeaderKey(header[i] || "", i)] = row[i] ?? "";
      }
      return obj;
    });

  if (!linhas.length) throw new Error("Nenhuma linha valida encontrada na aba Trechos.");
  return linhas;
}

function validatePeriodo(regiao: number, ano: number, mes: number) {
  if (!Number.isInteger(regiao) || regiao < 1 || regiao > 99) {
    throw new Error("Regiao invalida. Use valor inteiro entre 1 e 99.");
  }

  if (!Number.isInteger(ano) || ano < 2000 || ano > 2100) {
    throw new Error("Ano invalido. Use valor entre 2000 e 2100.");
  }

  if (!Number.isInteger(mes) || mes < 1 || mes > 12) {
    throw new Error("Mes invalido. Use valor entre 1 e 12.");
  }
}

type ImportarTrechosInput = {
  tipoFonte: "PAV" | "NAO_PAV";
  regiao: number;
  ano: number;
  mes: number;
  arquivoOrigem: string;
  linhas: unknown[];
  limparAntes?: boolean;
  dryRun?: boolean;
  operador?: string;
  perfil?: "OPERADOR" | "GESTOR" | "ADMIN";
};

async function registrarAuditoria(
  ctx: any,
  payload: {
    acao: string;
    entidade: string;
    regiao?: number;
    ano?: number;
    mes?: number;
    competencia?: string;
    operador?: string;
    perfil?: string;
    detalhes?: string;
  },
) {
  await ctx.db.insert("auditoriaEventos", {
    ...payload,
    criadoEm: Date.now(),
  });
}

async function upsertCompetencia(ctx: any, regiao: number, ano: number, mes: number) {
  const competencia = buildCompetencia(ano, mes);
  const now = Date.now();

  const competenciaExistente = await ctx.db
    .query("competencias")
    .withIndex("by_regiao_ano_mes", (q: any) => q.eq("regiao", regiao).eq("ano", ano).eq("mes", mes))
    .unique();

  if (competenciaExistente) {
    await ctx.db.patch(competenciaExistente._id, { atualizadoEm: now });
  } else {
    await ctx.db.insert("competencias", {
      regiao,
      ano,
      mes,
      competencia,
      criadoEm: now,
      atualizadoEm: now,
    });
  }
}

async function processarLinhasDaImportacao(
  ctx: any,
  params: {
    importacaoId: string;
    regiao: number;
    ano: number;
    mes: number;
    competencia: string;
    tipoFonte: TipoFonte;
    arquivoOrigem: string;
    dryRun: boolean;
    limparAntes: boolean;
    linhas: unknown[];
    operador?: string;
    perfil?: string;
  },
) {
  try {
    if (params.limparAntes && !params.dryRun) {
      const antigos = await ctx.db
        .query("trechos")
        .withIndex("by_regiao_ano_mes_tipo", (q: any) =>
          q
            .eq("regiao", params.regiao)
            .eq("ano", params.ano)
            .eq("mes", params.mes)
            .eq("tipoFonte", params.tipoFonte),
        )
        .collect();

      for (const doc of antigos) {
        await ctx.db.delete(doc._id);
      }
    }

    let linhasValidas = 0;
    let linhasIgnoradas = 0;
    let linhasComErro = 0;
    let gravados = 0;

    for (let i = 0; i < params.linhas.length; i += 1) {
      const row = (params.linhas[i] ?? {}) as Record<string, unknown>;
      const linhaPlanilha = i + 2;

      const trecho = toTrimmedString(getFirst(row, ["TRECHO", "TRECHOS", "trecho", "trechos"]));
      if (!trecho || normalizeKey(trecho) === "trecho") {
        linhasIgnoradas += 1;
        continue;
      }

      const sre = toTrimmedString(getFirst(row, ["S.R.E", "S.R.E.", "SRE", "sre"]));
      if (!sre) {
        linhasComErro += 1;
        await ctx.db.insert("importacaoErros", {
          importacaoId: params.importacaoId,
          regiao: params.regiao,
          ano: params.ano,
          mes: params.mes,
          competencia: params.competencia,
          tipoFonte: params.tipoFonte,
          linhaPlanilha,
          codigo: "SRE_OBRIGATORIO",
          mensagem: "Linha sem valor de S.R.E.",
          coluna: "S.R.E",
          criadoEm: Date.now(),
        });
        continue;
      }

      const extKmRaw = getFirst(row, ["EXT. (KM)", "EXT.(KM)", "EXT(KM)", "extKm", "EXT_KM"]);
      const extKm = toNumber(extKmRaw);
      if (toTrimmedString(extKmRaw) && extKm === undefined) {
        linhasComErro += 1;
        await ctx.db.insert("importacaoErros", {
          importacaoId: params.importacaoId,
          regiao: params.regiao,
          ano: params.ano,
          mes: params.mes,
          competencia: params.competencia,
          tipoFonte: params.tipoFonte,
          linhaPlanilha,
          codigo: "EXT_KM_INVALIDO",
          mensagem: "Valor de EXT. (KM) invalido.",
          coluna: "EXT. (KM)",
          valor: toTrimmedString(extKmRaw),
          criadoEm: Date.now(),
        });
        continue;
      }

      const regiaoConservacao = toTrimmedString(
        getFirst(row, [
          "REGIÃO DE CONSERVAÇÃO",
          "REGIAO DE CONSERVACAO",
          "regiaoConservacao",
          "REGIAO_CONSERVACAO",
        ]),
      );

      if (regiaoConservacao && Number(regiaoConservacao) !== params.regiao) {
        linhasIgnoradas += 1;
        continue;
      }

      const lote = toTrimmedString(getFirst(row, ["LOTE", "lote"]));
      const numero = toNumber(getFirst(row, ["N.", "N°", "numero", "N", "NUMERO"]));
      const cidadeSede = toTrimmedString(getFirst(row, ["CIDADE SEDE", "cidadeSede", "CIDADE_SEDE"]));
      const subtrechos = toTrimmedString(
        getFirst(row, ["SUBTRECHOS", "SUBTRECHO", "SBUTRECHO", "SEGMENTOS", "subtrechos"]),
      );
      const tipo = toTrimmedString(getFirst(row, ["TIPO", "tipo"]));
      const programacao = buildProgramacao(row, params.ano);

      linhasValidas += 1;

      if (!params.dryRun) {
        await ctx.db.insert("trechos", {
          regiao: params.regiao,
          ano: params.ano,
          mes: params.mes,
          competencia: params.competencia,
          tipoFonte: params.tipoFonte,
          lote,
          numero,
          regiaoConservacao,
          cidadeSede,
          trecho,
          sre,
          subtrechos,
          extKm,
          tipo,
          programacao,
          linhaPlanilha,
          importacaoId: params.importacaoId,
          arquivoOrigem: params.arquivoOrigem,
          importadoEm: Date.now(),
        });
        gravados += 1;
      }
    }

    await ctx.db.patch(params.importacaoId, {
      status: linhasComErro > 0 ? "SUCESSO_COM_ERROS" : "SUCESSO",
      totalLinhasRecebidas: params.linhas.length,
      linhasValidas,
      linhasIgnoradas,
      linhasComErro,
      gravados,
      finalizadoEm: Date.now(),
    });

    await registrarAuditoria(ctx, {
      acao: "IMPORTACAO_FINALIZADA",
      entidade: "trechos",
      regiao: params.regiao,
      ano: params.ano,
      mes: params.mes,
      competencia: params.competencia,
      operador: params.operador,
      perfil: params.perfil,
      detalhes: `${params.arquivoOrigem} (${params.tipoFonte}) - gravados=${gravados}, erros=${linhasComErro}`,
    });

    return {
      ok: true,
      importacaoId: params.importacaoId,
      competencia: params.competencia,
      regiao: params.regiao,
      ano: params.ano,
      mes: params.mes,
      tipoFonte: params.tipoFonte,
      arquivoOrigem: params.arquivoOrigem,
      dryRun: params.dryRun,
      linhasRecebidas: params.linhas.length,
      linhasValidas,
      linhasIgnoradas,
      linhasComErro,
      gravados,
    };
  } catch (error) {
    const erroFatal = error instanceof Error ? error.message : String(error);
    await ctx.db.patch(params.importacaoId, {
      status: "ERRO",
      erroFatal,
      finalizadoEm: Date.now(),
    });
    await registrarAuditoria(ctx, {
      acao: "IMPORTACAO_ERRO",
      entidade: "trechos",
      regiao: params.regiao,
      ano: params.ano,
      mes: params.mes,
      competencia: params.competencia,
      operador: params.operador,
      perfil: params.perfil,
      detalhes: erroFatal,
    });
    throw error;
  }
}

async function executarImportacao(ctx: any, args: ImportarTrechosInput) {
  const tipoFonte = args.tipoFonte as TipoFonte;
  const dryRun = args.dryRun === true;
  const limparAntes = args.limparAntes === true;
  const competencia = buildCompetencia(args.ano, args.mes);
  const iniciadoEm = Date.now();

  const importacaoId = await ctx.db.insert("importacoes", {
    regiao: args.regiao,
    ano: args.ano,
    mes: args.mes,
    competencia,
    tipoFonte,
    arquivoOrigem: args.arquivoOrigem,
    dryRun,
    limparAntes,
    status: "PROCESSANDO",
    totalLinhasRecebidas: args.linhas.length,
    linhasValidas: 0,
    linhasIgnoradas: 0,
    linhasComErro: 0,
    gravados: 0,
    iniciadoEm,
    operador: args.operador,
    perfil: args.perfil,
  });

  await registrarAuditoria(ctx, {
    acao: "IMPORTACAO_INICIADA",
    entidade: "trechos",
    regiao: args.regiao,
    ano: args.ano,
    mes: args.mes,
    competencia,
    operador: args.operador,
    perfil: args.perfil,
    detalhes: `${args.arquivoOrigem} (${tipoFonte})`,
  });

  await upsertCompetencia(ctx, args.regiao, args.ano, args.mes);
  return processarLinhasDaImportacao(ctx, {
    importacaoId,
    regiao: args.regiao,
    ano: args.ano,
    mes: args.mes,
    competencia,
    tipoFonte,
    arquivoOrigem: args.arquivoOrigem,
    dryRun,
    limparAntes,
    linhas: args.linhas,
    operador: args.operador,
    perfil: args.perfil,
  });
}

async function enfileirarImportacaoArquivo(
  ctx: any,
  args: {
    tipoFonte: TipoFonte;
    regiao: number;
    ano: number;
    mes: number;
    arquivoOrigem: string;
    arquivoBase64: string;
    limparAntes?: boolean;
    dryRun?: boolean;
    operador?: string;
    perfil?: "OPERADOR" | "GESTOR" | "ADMIN";
  },
) {
  const competencia = buildCompetencia(args.ano, args.mes);
  const iniciadoEm = Date.now();

  await upsertCompetencia(ctx, args.regiao, args.ano, args.mes);

  const importacaoId = await ctx.db.insert("importacoes", {
    regiao: args.regiao,
    ano: args.ano,
    mes: args.mes,
    competencia,
    tipoFonte: args.tipoFonte,
    arquivoOrigem: args.arquivoOrigem,
    dryRun: args.dryRun === true,
    limparAntes: args.limparAntes === true,
    status: "PROCESSANDO",
    totalLinhasRecebidas: 0,
    linhasValidas: 0,
    linhasIgnoradas: 0,
    linhasComErro: 0,
    gravados: 0,
    iniciadoEm,
    operador: args.operador,
    perfil: args.perfil,
  });

  await ctx.scheduler.runAfter(0, internal.trechos.processarImportacaoArquivoAsync, {
    importacaoId,
    arquivoBase64: args.arquivoBase64,
  });

  return importacaoId;
}

export const importarTrechos = mutation({
  args: {
    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    arquivoOrigem: v.string(),
    linhas: v.array(v.any()),
    limparAntes: v.optional(v.boolean()),
    dryRun: v.optional(v.boolean()),
    operador: v.optional(v.string()),
    perfil: v.optional(v.union(v.literal("OPERADOR"), v.literal("GESTOR"), v.literal("ADMIN"))),
  },
  handler: async (ctx, args) => {
    validatePeriodo(args.regiao, args.ano, args.mes);

    return executarImportacao(ctx, args);
  },
});

export const importarTrechosArquivo = mutation({
  args: {
    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    arquivoOrigem: v.string(),
    arquivoBase64: v.string(),
    limparAntes: v.optional(v.boolean()),
    dryRun: v.optional(v.boolean()),
    operador: v.optional(v.string()),
    perfil: v.optional(v.union(v.literal("OPERADOR"), v.literal("GESTOR"), v.literal("ADMIN"))),
  },
  handler: async (ctx, args) => {
    validatePeriodo(args.regiao, args.ano, args.mes);
    validateUploadSize(args.arquivoBase64);

    const linhas = lerLinhasTrechosDeArquivoBase64(args.arquivoBase64);
    return executarImportacao(ctx, {
      tipoFonte: args.tipoFonte,
      regiao: args.regiao,
      ano: args.ano,
      mes: args.mes,
      arquivoOrigem: args.arquivoOrigem,
      linhas,
      limparAntes: args.limparAntes,
      dryRun: args.dryRun,
      operador: args.operador,
      perfil: args.perfil,
    });
  },
});

export const iniciarImportacaoArquivoAsync = mutation({
  args: {
    sessionToken: v.string(),
    tipoFonte: v.union(v.literal("PAV"), v.literal("NAO_PAV")),
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    arquivoOrigem: v.string(),
    arquivoBase64: v.string(),
    limparAntes: v.optional(v.boolean()),
    dryRun: v.optional(v.boolean()),
    operador: v.optional(v.string()),
    perfil: v.optional(v.union(v.literal("OPERADOR"), v.literal("GESTOR"), v.literal("ADMIN"))),
  },
  handler: async (ctx, args) => {
    const sessao = await requireSession(ctx, args.sessionToken, ["OPERADOR", "GESTOR", "ADMIN"]);
    validatePeriodo(args.regiao, args.ano, args.mes);
    validateUploadSize(args.arquivoBase64);

    const importacaoId = await enfileirarImportacaoArquivo(ctx, {
      tipoFonte: args.tipoFonte,
      regiao: args.regiao,
      ano: args.ano,
      mes: args.mes,
      arquivoOrigem: args.arquivoOrigem,
      arquivoBase64: args.arquivoBase64,
      limparAntes: args.limparAntes,
      dryRun: args.dryRun,
      operador: args.operador,
      perfil: args.perfil,
    });

    await registrarAuditoria(ctx, {
      acao: "IMPORTACAO_ENFILEIRADA",
      entidade: "trechos",
      regiao: args.regiao,
      ano: args.ano,
      mes: args.mes,
      competencia: buildCompetencia(args.ano, args.mes),
      operador: args.operador,
      perfil: args.perfil,
      detalhes: `${args.arquivoOrigem} (${args.tipoFonte}) por ${sessao.nome}`,
    });

    return {
      ok: true,
      importacaoId,
      status: "PROCESSANDO",
      mensagem: "Importacao enfileirada para processamento.",
    };
  },
});

export const processarImportacaoArquivoAsync = internalMutation({
  args: {
    importacaoId: v.id("importacoes"),
    arquivoBase64: v.string(),
  },
  handler: async (ctx, args) => {
    const importacao = await ctx.db.get(args.importacaoId);
    if (!importacao) return { ok: false, motivo: "IMPORTACAO_NAO_ENCONTRADA" };
    if (importacao.status !== "PROCESSANDO") return { ok: false, motivo: "STATUS_INVALIDO" };

    try {
      const linhas = lerLinhasTrechosDeArquivoBase64(args.arquivoBase64);
      return processarLinhasDaImportacao(ctx, {
        importacaoId: String(importacao._id),
        regiao: importacao.regiao,
        ano: importacao.ano,
        mes: importacao.mes,
        competencia: importacao.competencia,
        tipoFonte: importacao.tipoFonte,
        arquivoOrigem: importacao.arquivoOrigem,
        dryRun: importacao.dryRun,
        limparAntes: importacao.limparAntes,
        linhas,
        operador: importacao.operador,
        perfil: importacao.perfil,
      });
    } catch (error) {
      const erroFatal = error instanceof Error ? error.message : String(error);
      await ctx.db.patch(importacao._id, {
        status: "ERRO",
        erroFatal,
        finalizadoEm: Date.now(),
      });
      throw error;
    }
  },
});

export const obterStatusImportacao = query({
  args: {
    importacaoId: v.id("importacoes"),
  },
  handler: async (ctx, args) => {
    const imp = await ctx.db.get(args.importacaoId);
    if (!imp) throw new Error("Importacao nao encontrada.");

    return {
      id: imp._id,
      status: imp.status,
      tipoFonte: imp.tipoFonte,
      arquivoOrigem: imp.arquivoOrigem,
      totalLinhasRecebidas: imp.totalLinhasRecebidas,
      linhasValidas: imp.linhasValidas,
      linhasIgnoradas: imp.linhasIgnoradas,
      linhasComErro: imp.linhasComErro,
      gravados: imp.gravados,
      erroFatal: imp.erroFatal,
      iniciadoEm: imp.iniciadoEm,
      finalizadoEm: imp.finalizadoEm,
    };
  },
});

export const iniciarImportacaoLoteAsync = mutation({
  args: {
    sessionToken: v.string(),
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
    pavArquivoOrigem: v.string(),
    pavArquivoBase64: v.string(),
    naoPavArquivoOrigem: v.string(),
    naoPavArquivoBase64: v.string(),
    limparAntes: v.optional(v.boolean()),
    dryRun: v.optional(v.boolean()),
    operador: v.optional(v.string()),
    perfil: v.optional(v.union(v.literal("OPERADOR"), v.literal("GESTOR"), v.literal("ADMIN"))),
  },
  handler: async (ctx, args) => {
    const sessao = await requireSession(ctx, args.sessionToken, ["OPERADOR", "GESTOR", "ADMIN"]);
    validatePeriodo(args.regiao, args.ano, args.mes);
    validateUploadSize(args.pavArquivoBase64);
    validateUploadSize(args.naoPavArquivoBase64);

    const pavId = await enfileirarImportacaoArquivo(ctx, {
      tipoFonte: "PAV",
      regiao: args.regiao,
      ano: args.ano,
      mes: args.mes,
      arquivoOrigem: args.pavArquivoOrigem,
      arquivoBase64: args.pavArquivoBase64,
      limparAntes: args.limparAntes,
      dryRun: args.dryRun,
      operador: args.operador,
      perfil: args.perfil,
    });

    const naoPavId = await enfileirarImportacaoArquivo(ctx, {
      tipoFonte: "NAO_PAV",
      regiao: args.regiao,
      ano: args.ano,
      mes: args.mes,
      arquivoOrigem: args.naoPavArquivoOrigem,
      arquivoBase64: args.naoPavArquivoBase64,
      limparAntes: args.limparAntes,
      dryRun: args.dryRun,
      operador: args.operador,
      perfil: args.perfil,
    });

    await registrarAuditoria(ctx, {
      acao: "IMPORTACAO_LOTE_ENFILEIRADA",
      entidade: "trechos",
      regiao: args.regiao,
      ano: args.ano,
      mes: args.mes,
      competencia: buildCompetencia(args.ano, args.mes),
      operador: args.operador,
      perfil: args.perfil,
      detalhes: `PAV=${args.pavArquivoOrigem}; NAO_PAV=${args.naoPavArquivoOrigem}; usuario=${sessao.nome}`,
    });

    return {
      ok: true,
      importacoes: {
        PAV: pavId,
        NAO_PAV: naoPavId,
      },
    };
  },
});

export const listarAuditoriaRecente = query({
  args: {
    sessionToken: v.string(),
    limite: v.optional(v.number()),
  },
  handler: async (ctx, args) => {
    await requireSession(ctx, args.sessionToken, ["GESTOR", "ADMIN", "OPERADOR"]);
    const limite = Math.min(Math.max(args.limite ?? 50, 1), 200);
    const eventos = await ctx.db.query("auditoriaEventos").withIndex("by_criadoEm").order("desc").take(limite);
    return eventos.map((e) => ({
      id: e._id,
      acao: e.acao,
      entidade: e.entidade,
      competencia: e.competencia,
      regiao: e.regiao,
      operador: e.operador,
      perfil: e.perfil,
      detalhes: e.detalhes,
      criadoEm: e.criadoEm,
    }));
  },
});

export const obterSaudeOperacional = query({
  args: {
    sessionToken: v.string(),
  },
  handler: async (ctx, args) => {
    await requireSession(ctx, args.sessionToken, ["OPERADOR", "GESTOR", "ADMIN"]);
    const agora = Date.now();
    const janela = agora - 24 * 60 * 60 * 1000;

    const importacoes = await ctx.db.query("importacoes").collect();
    const ultimas24h = importacoes.filter((i) => i.iniciadoEm >= janela);
    const processando = ultimas24h.filter((i) => i.status === "PROCESSANDO").length;
    const erros = ultimas24h.filter((i) => i.status === "ERRO").length;
    const sucessoComErros = ultimas24h.filter((i) => i.status === "SUCESSO_COM_ERROS").length;

    return {
      periodoHoras: 24,
      totalImportacoes: ultimas24h.length,
      processando,
      erros,
      sucessoComErros,
      taxaSucesso:
        ultimas24h.length === 0
          ? 100
          : Number((((ultimas24h.length - erros) / ultimas24h.length) * 100).toFixed(2)),
    };
  },
});

export const obterBaseConsolidada = query({
  args: {
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
  },
  handler: async (ctx, args) => {
    validatePeriodo(args.regiao, args.ano, args.mes);
    const competencia = buildCompetencia(args.ano, args.mes);

    const trechos = await ctx.db
      .query("trechos")
      .withIndex("by_regiao_ano_mes", (q) => q.eq("regiao", args.regiao).eq("ano", args.ano).eq("mes", args.mes))
      .collect();

    const importacoes = await ctx.db
      .query("importacoes")
      .withIndex("by_regiao_ano_mes", (q) => q.eq("regiao", args.regiao).eq("ano", args.ano).eq("mes", args.mes))
      .collect();

    const resumo = {
      totalTrechos: trechos.length,
      totalExtKm: trechos.reduce((acc, t) => acc + (t.extKm ?? 0), 0),
      porTipoFonte: {
        PAV: trechos.filter((t) => t.tipoFonte === "PAV").length,
        NAO_PAV: trechos.filter((t) => t.tipoFonte === "NAO_PAV").length,
      },
      totalImportacoes: importacoes.length,
      importacoesComErro: importacoes.filter((i) => i.status === "ERRO" || i.status === "SUCESSO_COM_ERROS").length,
    };

    return {
      regiao: args.regiao,
      ano: args.ano,
      mes: args.mes,
      competencia,
      resumo,
      importacoes,
      trechos,
    };
  },
});

export const obterGraficosCompetencia = query({
  args: {
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
  },
  handler: async (ctx, args) => {
    validatePeriodo(args.regiao, args.ano, args.mes);
    const competencia = buildCompetencia(args.ano, args.mes);

    const trechos = await ctx.db
      .query("trechos")
      .withIndex("by_regiao_ano_mes", (q) => q.eq("regiao", args.regiao).eq("ano", args.ano).eq("mes", args.mes))
      .collect();

    const totalTrechos = trechos.length;
    const totalKm = trechos.reduce((acc, item) => acc + (item.extKm ?? 0), 0);

    const trechosPav = trechos.filter((item) => item.tipoFonte === "PAV");
    const trechosNaoPav = trechos.filter((item) => item.tipoFonte === "NAO_PAV");

    const porTipoFonte = [
      {
        tipoFonte: "PAV",
        totalTrechos: trechosPav.length,
        totalKm: trechosPav.reduce((acc, item) => acc + (item.extKm ?? 0), 0),
      },
      {
        tipoFonte: "NAO_PAV",
        totalTrechos: trechosNaoPav.length,
        totalKm: trechosNaoPav.reduce((acc, item) => acc + (item.extKm ?? 0), 0),
      },
    ];

    const kmPorSreMap = new Map<string, number>();
    for (const item of trechos) {
      const key = item.sre ?? "SEM_SRE";
      kmPorSreMap.set(key, (kmPorSreMap.get(key) ?? 0) + (item.extKm ?? 0));
    }

    const topSrePorKm = Array.from(kmPorSreMap.entries())
      .map(([sre, km]) => ({ sre, km }))
      .sort((a, b) => b.km - a.km)
      .slice(0, 10);

    const porTipoViaMap = new Map<string, number>();
    for (const item of trechos) {
      const key = item.tipo ?? "SEM_TIPO";
      porTipoViaMap.set(key, (porTipoViaMap.get(key) ?? 0) + 1);
    }

    const porTipoVia = Array.from(porTipoViaMap.entries())
      .map(([tipo, totalTrechosTipo]) => ({ tipo, totalTrechos: totalTrechosTipo }))
      .sort((a, b) => b.totalTrechos - a.totalTrechos);

    const programadosNoMes = trechos.filter((item) => item.programacao[competencia] === true).length;
    const naoProgramadosNoMes = totalTrechos - programadosNoMes;

    return {
      regiao: args.regiao,
      ano: args.ano,
      mes: args.mes,
      competencia,
      kpis: {
        totalTrechos,
        totalKm,
        programadosNoMes,
        naoProgramadosNoMes,
      },
      series: {
        porTipoFonte,
        topSrePorKm,
        porTipoVia,
      },
    };
  },
});

export const obterEvolucaoManutencao = query({
  args: {
    regiao: v.number(),
    ano: v.number(),
    trecho: v.optional(v.string()),
  },
  handler: async (ctx, args) => {
    if (!Number.isInteger(args.regiao) || args.regiao < 1 || args.regiao > 99) {
      throw new Error("Regiao invalida. Use valor inteiro entre 1 e 99.");
    }
    if (!Number.isInteger(args.ano) || args.ano < 2000 || args.ano > 2100) {
      throw new Error("Ano invalido. Use valor entre 2000 e 2100.");
    }

    const todos = await ctx.db.query("trechos").collect();
    const trechosAno = todos.filter((t) => t.regiao === args.regiao && t.ano === args.ano);

    const snapshotsMaisRecentes = new Map<string, (typeof trechosAno)[number]>();
    for (const item of trechosAno) {
      const key = [
        item.tipoFonte,
        item.trecho,
        item.sre ?? "",
        item.subtrechos ?? "",
        item.tipo ?? "",
        item.extKm ?? "",
      ].join("|");

      const existente = snapshotsMaisRecentes.get(key);
      if (!existente) {
        snapshotsMaisRecentes.set(key, item);
        continue;
      }

      if (item.mes > existente.mes || (item.mes === existente.mes && item.importadoEm > existente.importadoEm)) {
        snapshotsMaisRecentes.set(key, item);
      }
    }

    const baseEvolucao = Array.from(snapshotsMaisRecentes.values());
    const snapshotsPorMes = new Map<number, typeof baseEvolucao>();
    for (const mes of Array.from({ length: 12 }, (_, i) => i + 1)) {
      const itensMes = trechosAno.filter((t) => t.mes === mes);
      const recentesMes = new Map<string, (typeof trechosAno)[number]>();

      for (const item of itensMes) {
        const key = [
          item.tipoFonte,
          item.trecho,
          item.sre ?? "",
          item.subtrechos ?? "",
          item.tipo ?? "",
          item.extKm ?? "",
        ].join("|");

        const existente = recentesMes.get(key);
        if (!existente || item.importadoEm > existente.importadoEm) {
          recentesMes.set(key, item);
        }
      }

      snapshotsPorMes.set(mes, Array.from(recentesMes.values()));
    }

    const trechosDisponiveis = Array.from(new Set(baseEvolucao.map((t) => t.trecho).filter(Boolean))).sort((a, b) =>
      a.localeCompare(b, "pt-BR"),
    );

    const trechoSelecionado = args.trecho && trechosDisponiveis.includes(args.trecho) ? args.trecho : trechosDisponiveis[0] ?? "";
    const meses = Array.from({ length: 12 }, (_, i) => i + 1);

    const mensal = meses.map((mes) => {
      const competencia = `${args.ano}-${String(mes).padStart(2, "0")}`;
      let geralKm = 0;
      let trechoKm = 0;
      let geralKmCarregado = 0;
      let trechoKmCarregado = 0;

      const baseMes = snapshotsPorMes.get(mes) ?? [];
      for (const item of baseMes) {
        const km = item.extKm ?? 0;
        geralKmCarregado += km;
        if (trechoSelecionado && item.trecho === trechoSelecionado) trechoKmCarregado += km;
      }

      for (const item of baseEvolucao) {
        if (item.programacao?.[competencia] !== true) continue;
        const km = item.extKm ?? 0;
        geralKm += km;
        if (trechoSelecionado && item.trecho === trechoSelecionado) trechoKm += km;
      }

      return {
        mes,
        competencia,
        geralKm: Number(geralKm.toFixed(2)),
        trechoKm: Number(trechoKm.toFixed(2)),
        geralKmCarregado: Number(geralKmCarregado.toFixed(2)),
        trechoKmCarregado: Number(trechoKmCarregado.toFixed(2)),
      };
    });

    return {
      regiao: args.regiao,
      ano: args.ano,
      trechoSelecionado,
      trechosDisponiveis,
      mensal,
    };
  },
});

export const obterInconsistenciasImportacao = query({
  args: {
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
  },
  handler: async (ctx, args) => {
    validatePeriodo(args.regiao, args.ano, args.mes);
    const competencia = buildCompetencia(args.ano, args.mes);

    const importacoes = await ctx.db
      .query("importacoes")
      .withIndex("by_regiao_ano_mes", (q) => q.eq("regiao", args.regiao).eq("ano", args.ano).eq("mes", args.mes))
      .collect();

    const erros: Array<{
      importacaoId: string;
      tipoFonte: TipoFonte;
      linhaPlanilha: number;
      codigo: string;
      mensagem: string;
      coluna?: string;
      valor?: string;
    }> = [];

    for (const imp of importacoes) {
      const errosImportacao = await ctx.db
        .query("importacaoErros")
        .withIndex("by_importacao", (q) => q.eq("importacaoId", imp._id))
        .collect();

      for (const e of errosImportacao) {
        erros.push({
          importacaoId: String(imp._id),
          tipoFonte: imp.tipoFonte,
          linhaPlanilha: e.linhaPlanilha,
          codigo: e.codigo,
          mensagem: e.mensagem,
          coluna: e.coluna,
          valor: e.valor,
        });
      }
    }

    const porCodigoMap = new Map<string, number>();
    for (const e of erros) {
      porCodigoMap.set(e.codigo, (porCodigoMap.get(e.codigo) ?? 0) + 1);
    }

    const porCodigo = Array.from(porCodigoMap.entries())
      .map(([codigo, total]) => ({ codigo, total }))
      .sort((a, b) => b.total - a.total);

    return {
      regiao: args.regiao,
      ano: args.ano,
      mes: args.mes,
      competencia,
      resumo: {
        totalImportacoes: importacoes.length,
        importacoesComErro: importacoes.filter((i) => i.status === "ERRO" || i.status === "SUCESSO_COM_ERROS").length,
        totalErros: erros.length,
      },
      porCodigo,
      erros: erros.slice(0, 200),
      importacoes: importacoes.map((i) => ({
        id: i._id,
        tipoFonte: i.tipoFonte,
        arquivoOrigem: i.arquivoOrigem,
        status: i.status,
        iniciadoEm: i.iniciadoEm,
        finalizadoEm: i.finalizadoEm,
        linhasComErro: i.linhasComErro,
        linhasIgnoradas: i.linhasIgnoradas,
        linhasValidas: i.linhasValidas,
        gravados: i.gravados,
        dryRun: i.dryRun,
      })),
    };
  },
});

export const gerarPayloadRelatorio = query({
  args: {
    regiao: v.number(),
    ano: v.number(),
    mes: v.number(),
  },
  handler: async (ctx, args) => {
    validatePeriodo(args.regiao, args.ano, args.mes);
    const competencia = buildCompetencia(args.ano, args.mes);

    const base = await ctx.db
      .query("trechos")
      .withIndex("by_regiao_ano_mes", (q) => q.eq("regiao", args.regiao).eq("ano", args.ano).eq("mes", args.mes))
      .collect();

    const graficos = await (async () => {
      const totalTrechos = base.length;
      const totalKm = base.reduce((acc, item) => acc + (item.extKm ?? 0), 0);
      const programadosNoMes = base.filter((item) => item.programacao[competencia] === true).length;

      return {
        kpis: {
          totalTrechos,
          totalKm,
          programadosNoMes,
          percentualProgramados: totalTrechos === 0 ? 0 : Number(((programadosNoMes / totalTrechos) * 100).toFixed(2)),
        },
      };
    })();

    const inconsistencias = await (async () => {
      const importacoes = await ctx.db
        .query("importacoes")
        .withIndex("by_regiao_ano_mes", (q) => q.eq("regiao", args.regiao).eq("ano", args.ano).eq("mes", args.mes))
        .collect();

      const totalErros = importacoes.reduce((acc, item) => acc + item.linhasComErro, 0);
      return {
        totalImportacoes: importacoes.length,
        totalErros,
      };
    })();

    return {
      metadata: {
        regiao: args.regiao,
        ano: args.ano,
        mes: args.mes,
        competencia,
        geradoEm: Date.now(),
      },
      graficos,
      inconsistencias,
      observacoesAutomaticas: [
        `Competencia analisada: ${competencia} (Regiao ${args.regiao}).`,
        `Total de trechos no periodo: ${base.length}.`,
        `Total de importacoes com rastreabilidade: ${inconsistencias.totalImportacoes}.`,
      ],
    };
  },
});
