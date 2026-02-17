import { v } from "convex/values";
import { mutation, query } from "./_generated/server";

type Perfil = "OPERADOR" | "GESTOR" | "ADMIN";

function normalizeEmail(email: string): string {
  return email.trim().toLowerCase();
}

async function sha256(text: string): Promise<string> {
  const data = new TextEncoder().encode(text);
  const hash = await crypto.subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(hash))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

function newToken(): string {
  return `${crypto.randomUUID()}-${Date.now()}`;
}

function validarPoliticaSenha(senha: string) {
  if (senha.length < 8) throw new Error("A senha deve ter ao menos 8 caracteres.");
  if (!/[a-z]/.test(senha)) throw new Error("A senha deve conter letra minuscula.");
  if (!/[A-Z]/.test(senha)) throw new Error("A senha deve conter letra maiuscula.");
  if (!/[0-9]/.test(senha)) throw new Error("A senha deve conter numero.");
  if (!/[^a-zA-Z0-9]/.test(senha)) throw new Error("A senha deve conter caractere especial.");
}

export const login = mutation({
  args: {
    email: v.string(),
    senha: v.string(),
  },
  handler: async (ctx, args) => {
    const email = normalizeEmail(args.email);
    const senhaHash = await sha256(args.senha);

    const totalUsuarios = (await ctx.db.query("usuarios").take(1)).length;

    let usuario = await ctx.db.query("usuarios").withIndex("by_email", (q) => q.eq("email", email)).unique();

    if (!usuario && totalUsuarios === 0) {
      validarPoliticaSenha(args.senha);
      const now = Date.now();
      const userId = await ctx.db.insert("usuarios", {
        nome: "Administrador",
        email,
        senhaHash,
        perfil: "ADMIN",
        ativo: true,
        forcarTrocaSenha: true,
        criadoEm: now,
        atualizadoEm: now,
      });
      usuario = await ctx.db.get(userId);
    }

    if (!usuario || !usuario.ativo) throw new Error("Usuario ou senha invalido.");
    if (usuario.senhaHash !== senhaHash) throw new Error("Usuario ou senha invalido.");

    const now = Date.now();
    const expiraEm = now + 1000 * 60 * 60 * 12;
    const token = newToken();

    await ctx.db.insert("sessoes", {
      usuarioId: usuario._id,
      token,
      nome: usuario.nome,
      email: usuario.email,
      perfil: usuario.perfil,
      criadoEm: now,
      expiraEm,
    });

    return {
      token,
      usuario: {
        id: usuario._id,
        nome: usuario.nome,
        email: usuario.email,
        perfil: usuario.perfil,
      },
      precisaTrocaSenha: usuario.forcarTrocaSenha === true,
      expiraEm,
    };
  },
});

export const me = query({
  args: { sessionToken: v.string() },
  handler: async (ctx, args) => {
    const sessao = await ctx.db.query("sessoes").withIndex("by_token", (q) => q.eq("token", args.sessionToken)).unique();
    if (!sessao || sessao.revogadaEm || sessao.expiraEm < Date.now()) return null;
    const usuario = await ctx.db.get(sessao.usuarioId);

    return {
      token: sessao.token,
      usuario: {
        id: sessao.usuarioId,
        nome: sessao.nome,
        email: sessao.email,
        perfil: sessao.perfil,
      },
      precisaTrocaSenha: usuario?.forcarTrocaSenha === true,
      expiraEm: sessao.expiraEm,
    };
  },
});

export const logout = mutation({
  args: { sessionToken: v.string() },
  handler: async (ctx, args) => {
    const sessao = await ctx.db.query("sessoes").withIndex("by_token", (q) => q.eq("token", args.sessionToken)).unique();
    if (!sessao || sessao.revogadaEm) return { ok: true };

    await ctx.db.patch(sessao._id, { revogadaEm: Date.now() });
    return { ok: true };
  },
});

export const criarUsuario = mutation({
  args: {
    sessionToken: v.string(),
    nome: v.string(),
    email: v.string(),
    senha: v.string(),
    perfil: v.union(v.literal("OPERADOR"), v.literal("GESTOR"), v.literal("ADMIN")),
  },
  handler: async (ctx, args) => {
    const sessao = await ctx.db.query("sessoes").withIndex("by_token", (q) => q.eq("token", args.sessionToken)).unique();
    if (!sessao || sessao.revogadaEm || sessao.expiraEm < Date.now()) throw new Error("Sessao invalida.");
    if (sessao.perfil !== "ADMIN") throw new Error("Apenas ADMIN pode criar usuario.");

    const email = normalizeEmail(args.email);
    const existente = await ctx.db.query("usuarios").withIndex("by_email", (q) => q.eq("email", email)).unique();
    if (existente) throw new Error("Email ja cadastrado.");

    validarPoliticaSenha(args.senha);

    const now = Date.now();
    await ctx.db.insert("usuarios", {
      nome: args.nome.trim(),
      email,
      senhaHash: await sha256(args.senha),
      perfil: args.perfil as Perfil,
      ativo: true,
      forcarTrocaSenha: true,
      criadoEm: now,
      atualizadoEm: now,
    });

    return { ok: true };
  },
});

export const listarUsuarios = query({
  args: {
    sessionToken: v.string(),
  },
  handler: async (ctx, args) => {
    const sessao = await ctx.db.query("sessoes").withIndex("by_token", (q) => q.eq("token", args.sessionToken)).unique();
    if (!sessao || sessao.revogadaEm || sessao.expiraEm < Date.now()) throw new Error("Sessao invalida.");
    if (sessao.perfil !== "ADMIN") throw new Error("Apenas ADMIN pode listar usuarios.");

    const usuarios = await ctx.db.query("usuarios").collect();
    return usuarios
      .map((u) => ({
        id: u._id,
        nome: u.nome,
        email: u.email,
        perfil: u.perfil,
        ativo: u.ativo,
        forcarTrocaSenha: u.forcarTrocaSenha === true,
        atualizadoEm: u.atualizadoEm,
      }))
      .sort((a, b) => b.atualizadoEm - a.atualizadoEm);
  },
});

export const atualizarUsuario = mutation({
  args: {
    sessionToken: v.string(),
    usuarioId: v.id("usuarios"),
    nome: v.optional(v.string()),
    perfil: v.optional(v.union(v.literal("OPERADOR"), v.literal("GESTOR"), v.literal("ADMIN"))),
    ativo: v.optional(v.boolean()),
  },
  handler: async (ctx, args) => {
    const sessao = await ctx.db.query("sessoes").withIndex("by_token", (q) => q.eq("token", args.sessionToken)).unique();
    if (!sessao || sessao.revogadaEm || sessao.expiraEm < Date.now()) throw new Error("Sessao invalida.");
    if (sessao.perfil !== "ADMIN") throw new Error("Apenas ADMIN pode atualizar usuarios.");

    const usuario = await ctx.db.get(args.usuarioId);
    if (!usuario) throw new Error("Usuario nao encontrado.");

    await ctx.db.patch(args.usuarioId, {
      nome: args.nome?.trim() || usuario.nome,
      perfil: args.perfil ?? usuario.perfil,
      ativo: args.ativo ?? usuario.ativo,
      atualizadoEm: Date.now(),
    });

    return { ok: true };
  },
});

export const redefinirSenhaUsuario = mutation({
  args: {
    sessionToken: v.string(),
    usuarioId: v.id("usuarios"),
    novaSenhaTemporaria: v.string(),
  },
  handler: async (ctx, args) => {
    const sessao = await ctx.db.query("sessoes").withIndex("by_token", (q) => q.eq("token", args.sessionToken)).unique();
    if (!sessao || sessao.revogadaEm || sessao.expiraEm < Date.now()) throw new Error("Sessao invalida.");
    if (sessao.perfil !== "ADMIN") throw new Error("Apenas ADMIN pode redefinir senha.");

    const usuario = await ctx.db.get(args.usuarioId);
    if (!usuario) throw new Error("Usuario nao encontrado.");

    validarPoliticaSenha(args.novaSenhaTemporaria);

    await ctx.db.patch(args.usuarioId, {
      senhaHash: await sha256(args.novaSenhaTemporaria),
      forcarTrocaSenha: true,
      atualizadoEm: Date.now(),
    });

    return { ok: true };
  },
});

export const alterarSenha = mutation({
  args: {
    sessionToken: v.string(),
    senhaAtual: v.string(),
    novaSenha: v.string(),
  },
  handler: async (ctx, args) => {
    const sessao = await ctx.db.query("sessoes").withIndex("by_token", (q) => q.eq("token", args.sessionToken)).unique();
    if (!sessao || sessao.revogadaEm || sessao.expiraEm < Date.now()) throw new Error("Sessao invalida.");

    const usuario = await ctx.db.get(sessao.usuarioId);
    if (!usuario || !usuario.ativo) throw new Error("Usuario invalido.");

    const senhaAtualHash = await sha256(args.senhaAtual);
    if (senhaAtualHash !== usuario.senhaHash) throw new Error("Senha atual incorreta.");

    validarPoliticaSenha(args.novaSenha);

    await ctx.db.patch(usuario._id, {
      senhaHash: await sha256(args.novaSenha),
      forcarTrocaSenha: false,
      senhaAtualizadaEm: Date.now(),
      atualizadoEm: Date.now(),
    });

    return { ok: true };
  },
});
