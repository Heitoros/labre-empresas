import { CommonModule } from "@angular/common";
import { Component, OnInit } from "@angular/core";
import { FormsModule } from "@angular/forms";
import { ConvexHttpClient } from "convex/browser";
import {
  AlignmentType,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableOfContents,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import JSZip from "jszip";
import { api } from "../../convex/_generated/api";
import { environment } from "../environments/environment";

type FonteResumo = { tipoFonte: string; totalTrechos: number; totalKm: number };
type SreResumo = { sre: string; km: number };
type CodigoInconsistenciaResumo = { codigo: string; total: number };
type ImportacaoResumo = {
  id: string;
  tipoFonte: string;
  status: string;
  arquivoOrigem: string;
  iniciadoEm?: number;
  finalizadoEm?: number;
  linhasValidas: number;
  linhasIgnoradas: number;
  linhasComErro: number;
  gravados: number;
};
type AuditoriaEvento = {
  id: string;
  acao: string;
  entidade: string;
  competencia?: string;
  regiao?: number;
  operador?: string;
  perfil?: string;
  detalhes?: string;
  criadoEm: number;
};
type ImportacaoStatus = {
  id: string;
  status: string;
  gravados: number;
  linhasIgnoradas: number;
  linhasComErro: number;
  erroFatal?: string;
};
type SessaoAtual = {
  token: string;
  usuario: {
    id: string;
    nome: string;
    email: string;
    perfil: "OPERADOR" | "GESTOR" | "ADMIN";
  };
  expiraEm: number;
  precisaTrocaSenha?: boolean;
};
type UsuarioAdmin = {
  id: string;
  nome: string;
  email: string;
  perfil: "OPERADOR" | "GESTOR" | "ADMIN";
  ativo: boolean;
  forcarTrocaSenha: boolean;
  atualizadoEm: number;
};
type SaudeOperacional = {
  periodoHoras: number;
  totalImportacoes: number;
  processando: number;
  erros: number;
  sucessoComErros: number;
  taxaSucesso: number;
};
type WorkbookSerie = { label: string; valor: number; percentual: number };
type WorkbookGrafico = {
  id: string;
  aba: string;
  ordem: number;
  titulo: string;
  tipoGrafico: string;
  trecho?: string;
  total: number;
  series: WorkbookSerie[];
};
type WorkbookTrechoGroup = {
  trecho: string;
  graficos: WorkbookGrafico[];
  tituloTemplate?: string;
};
type EvolucaoManutencaoItem = {
  mes: number;
  competencia: string;
  geralKm: number;
  trechoKm: number;
};
type RelatorioTemplateHeading = {
  index: number;
  nivel: number;
  texto: string;
  origem: "estilo" | "numeracao";
};

@Component({
  selector: "app-root",
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: "./app.component.html",
  styleUrl: "./app.component.css",
})
export class AppComponent implements OnInit {
  regiao = 1;
  ano = 2025;
  mes = 12;
  loading = false;
  erro = "";

  totalTrechos = 0;
  totalKm = 0;
  programados = 0;
  naoProgramados = 0;

  porTipoFonte: FonteResumo[] = [];
  topSrePorKm: SreResumo[] = [];
  porCodigoInconsistencia: CodigoInconsistenciaResumo[] = [];
  importacoes: ImportacaoResumo[] = [];
  mostrarHistoricoCompleto = false;
  gerandoConsolidado = false;
  uploadTipoFonte: "PAV" | "NAO_PAV" = "PAV";
  uploadLimparAntes = true;
  uploadProcessarComplementar = true;
  operador = "";
  perfil: "OPERADOR" | "GESTOR" | "ADMIN" = "OPERADOR";
  emailLogin = "";
  senhaLogin = "";
  authMensagem = "";
  sessaoAtual: SessaoAtual | null = null;
  senhaAtualTroca = "";
  novaSenhaTroca = "";
  confirmarSenhaTroca = "";
  arquivoSelecionado: File | null = null;
  arquivoPavSelecionado: File | null = null;
  arquivoNaoPavSelecionado: File | null = null;
  uploadEmAndamento = false;
  uploadImportacaoId = "";
  uploadMensagem = "";
  auditoriaRecente: AuditoriaEvento[] = [];
  usuariosAdmin: UsuarioAdmin[] = [];
  saudeOperacional: SaudeOperacional | null = null;
  filtroUsuarios = "";
  novoUsuarioNome = "";
  novoUsuarioEmail = "";
  novoUsuarioSenha = "Temp@123";
  novoUsuarioPerfil: "OPERADOR" | "GESTOR" | "ADMIN" = "OPERADOR";
  confirmacaoModalAberto = false;
  confirmacaoModalTitulo = "";
  confirmacaoModalMensagem = "";
  confirmacaoAcao: "RESET_SENHA" | "ALTERAR_ATIVO" | null = null;
  confirmacaoUsuario: UsuarioAdmin | null = null;
  confirmacaoNovoAtivo: boolean | null = null;
  private sessionExpiryTimer: ReturnType<typeof setTimeout> | null = null;
  resumoInconsistencias = {
    totalImportacoes: 0,
    importacoesComErro: 0,
    totalErros: 0,
  };
  abaGraficos: "tipoFonte" | "programacao" = "tipoFonte";
  workbookTipoFonte: "PAV" | "NAO_PAV" = "PAV";
  workbookTrechoSelecionado = "";
  workbookGraficos: WorkbookGrafico[] = [];
  evolucaoTrechosDisponiveis: string[] = [];
  evolucaoTrechoSelecionado = "";
  evolucaoMensal: EvolucaoManutencaoItem[] = [];
  templateEstruturaNome = "";
  templateEstruturaMensagem = "";
  templateEstruturaHeadings: RelatorioTemplateHeading[] = [];
  templateDocxBase: File | null = null;
  templateDocxMensagem = "";
  relatorioLarguraBlocoCm = 16;
  relatorioAlturaPavCm = 15;
  relatorioAlturaNaoPavCm = 7;

  meses = [
    { value: 1, label: "Janeiro" },
    { value: 2, label: "Fevereiro" },
    { value: 3, label: "Marco" },
    { value: 4, label: "Abril" },
    { value: 5, label: "Maio" },
    { value: 6, label: "Junho" },
    { value: 7, label: "Julho" },
    { value: 8, label: "Agosto" },
    { value: 9, label: "Setembro" },
    { value: 10, label: "Outubro" },
    { value: 11, label: "Novembro" },
    { value: 12, label: "Dezembro" },
  ];

  regioes = [1, 2, 3, 11, 12, 13];

  private client = new ConvexHttpClient(environment.convexUrl);

  ngOnInit(): void {
    this.operador = localStorage.getItem("labre_operador") ?? "";
    this.perfil = (localStorage.getItem("labre_perfil") as "OPERADOR" | "GESTOR" | "ADMIN") ?? "OPERADOR";
    this.relatorioLarguraBlocoCm = Number(localStorage.getItem("labre_relatorio_largura_cm") ?? this.relatorioLarguraBlocoCm);
    this.relatorioAlturaPavCm = Number(localStorage.getItem("labre_relatorio_altura_pav_cm") ?? this.relatorioAlturaPavCm);
    this.relatorioAlturaNaoPavCm = Number(localStorage.getItem("labre_relatorio_altura_naopav_cm") ?? this.relatorioAlturaNaoPavCm);
    this.salvarDimensoesRelatorio();
    void this.inicializarSessao();
  }

  salvarDimensoesRelatorio(): void {
    const clamp = (value: number, min: number, max: number): number => {
      if (!Number.isFinite(value)) return min;
      return Math.max(min, Math.min(max, Number(value.toFixed(1))));
    };

    this.relatorioLarguraBlocoCm = clamp(this.relatorioLarguraBlocoCm, 8, 21);
    this.relatorioAlturaPavCm = clamp(this.relatorioAlturaPavCm, 6, 25);
    this.relatorioAlturaNaoPavCm = clamp(this.relatorioAlturaNaoPavCm, 4, 25);

    localStorage.setItem("labre_relatorio_largura_cm", String(this.relatorioLarguraBlocoCm));
    localStorage.setItem("labre_relatorio_altura_pav_cm", String(this.relatorioAlturaPavCm));
    localStorage.setItem("labre_relatorio_altura_naopav_cm", String(this.relatorioAlturaNaoPavCm));
  }

  private tokenSessao(): string {
    if (!this.sessaoAtual?.token) throw new Error("Sessao nao autenticada.");
    return this.sessaoAtual.token;
  }

  private async inicializarSessao(): Promise<void> {
    const token = localStorage.getItem("labre_session_token");
    if (!token) return;

    try {
      const sessao = await this.client.query(api.auth.me, { sessionToken: token });
      if (!sessao) {
        localStorage.removeItem("labre_session_token");
        return;
      }
      this.sessaoAtual = sessao as SessaoAtual;
      this.agendarExpiracaoSessao();
      if (!this.operador) this.operador = this.sessaoAtual.usuario.nome;
      this.perfil = this.sessaoAtual.usuario.perfil;
      this.salvarPerfilOperacao();
      if (!this.sessaoAtual.precisaTrocaSenha) {
        await this.recarregar();
      }
    } catch {
      localStorage.removeItem("labre_session_token");
    }
  }

  private agendarExpiracaoSessao(): void {
    if (this.sessionExpiryTimer) {
      clearTimeout(this.sessionExpiryTimer);
      this.sessionExpiryTimer = null;
    }
    if (!this.sessaoAtual?.expiraEm) return;

    const ms = this.sessaoAtual.expiraEm - Date.now();
    if (ms <= 0) {
      void this.logoutPorExpiracao();
      return;
    }

    this.sessionExpiryTimer = setTimeout(() => {
      void this.logoutPorExpiracao();
    }, ms + 1000);
  }

  private async logoutPorExpiracao(): Promise<void> {
    await this.logout();
    this.authMensagem = "Sessao expirada. Faca login novamente.";
  }

  private erroAutenticacao(message: string): boolean {
    const m = message.toLowerCase();
    return m.includes("sessao") || m.includes("sessão") || m.includes("expirada") || m.includes("autenticada");
  }

  private async tratarErroAutenticacao(e: unknown): Promise<boolean> {
    const msg = e instanceof Error ? e.message : String(e);
    if (!this.erroAutenticacao(msg)) return false;
    await this.logout();
    this.authMensagem = "Sua sessao expirou ou ficou invalida. Entre novamente.";
    return true;
  }

  private mensagemErroAuthAmigavel(e: unknown, contexto: "LOGIN" | "TROCA_SENHA" | "GERAL" = "GERAL"): string {
    const msg = e instanceof Error ? e.message : String(e);
    const m = msg.toLowerCase();

    if (m.includes("usuario ou senha invalido") || m.includes("usuário ou senha inválido")) {
      return "Email ou senha invalidos. Confira os dados e tente novamente.";
    }

    if (m.includes("senha atual incorreta")) {
      return "Senha atual incorreta. Tente novamente.";
    }

    if (m.includes("sessao invalida") || m.includes("sessão inválida") || m.includes("sessao expirada")) {
      return "Sua sessao expirou. Faca login novamente.";
    }

    if (m.includes("a senha deve")) {
      return msg;
    }

    if (m.includes("server error") || m.includes("[request id:")) {
      if (contexto === "LOGIN") {
        return "Nao foi possivel concluir o login agora. Verifique email/senha e tente novamente.";
      }
      if (contexto === "TROCA_SENHA") {
        return "Nao foi possivel salvar a nova senha agora. Confira a senha atual e tente novamente.";
      }
      return "Ocorreu um erro temporario no servidor. Tente novamente em instantes.";
    }

    return msg;
  }

  async login(): Promise<void> {
    this.authMensagem = "";
    try {
      const sessao = await this.client.mutation(api.auth.login, {
        email: this.emailLogin,
        senha: this.senhaLogin,
      });
      this.sessaoAtual = sessao as SessaoAtual;
      this.agendarExpiracaoSessao();
      localStorage.setItem("labre_session_token", this.sessaoAtual.token);
      this.operador = this.sessaoAtual.usuario.nome;
      this.perfil = this.sessaoAtual.usuario.perfil;
      this.salvarPerfilOperacao();
      this.authMensagem = "Login realizado com sucesso.";
      this.senhaLogin = "";
      if (!this.sessaoAtual.precisaTrocaSenha) {
        await this.recarregar();
      } else {
        this.authMensagem = "Troca de senha obrigatoria no primeiro acesso.";
      }
    } catch (e) {
      this.authMensagem = this.mensagemErroAuthAmigavel(e, "LOGIN");
    }
  }

  async logout(): Promise<void> {
    try {
      if (this.sessaoAtual?.token) {
        await this.client.mutation(api.auth.logout, { sessionToken: this.sessaoAtual.token });
      }
    } finally {
      this.sessaoAtual = null;
      this.usuariosAdmin = [];
      localStorage.removeItem("labre_session_token");
      if (this.sessionExpiryTimer) {
        clearTimeout(this.sessionExpiryTimer);
        this.sessionExpiryTimer = null;
      }
    }
  }

  async criarUsuarioAdmin(): Promise<void> {
    if (!this.sessaoAtual?.token) return;
    try {
      await this.client.mutation(api.auth.criarUsuario, {
        sessionToken: this.sessaoAtual.token,
        nome: this.novoUsuarioNome,
        email: this.novoUsuarioEmail,
        senha: this.novoUsuarioSenha,
        perfil: this.novoUsuarioPerfil,
      });
      this.authMensagem = "Usuario criado com sucesso.";
      this.novoUsuarioNome = "";
      this.novoUsuarioEmail = "";
      this.novoUsuarioSenha = "Temp@123";
      this.novoUsuarioPerfil = "OPERADOR";
      await this.carregarUsuariosAdmin();
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.authMensagem = e instanceof Error ? e.message : String(e);
    }
  }

  async atualizarUsuarioAdmin(user: UsuarioAdmin): Promise<void> {
    if (!this.sessaoAtual?.token) return;
    try {
      await this.client.mutation(api.auth.atualizarUsuario, {
        sessionToken: this.sessaoAtual.token,
        usuarioId: user.id as any,
        nome: user.nome,
        perfil: user.perfil,
        ativo: user.ativo,
      });
      await this.carregarUsuariosAdmin();
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.authMensagem = e instanceof Error ? e.message : String(e);
    }
  }

  async redefinirSenhaUsuarioAdmin(user: UsuarioAdmin): Promise<void> {
    if (!this.sessaoAtual?.token) return;
    try {
      await this.client.mutation(api.auth.redefinirSenhaUsuario, {
        sessionToken: this.sessaoAtual.token,
        usuarioId: user.id as any,
        novaSenhaTemporaria: "Temp@123",
      });
      this.authMensagem = `Senha temporaria redefinida para ${user.email}.`;
      await this.carregarUsuariosAdmin();
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.authMensagem = e instanceof Error ? e.message : String(e);
    }
  }

  private async carregarUsuariosAdmin(): Promise<void> {
    if (!this.sessaoAtual?.token || this.sessaoAtual.usuario.perfil !== "ADMIN") {
      this.usuariosAdmin = [];
      return;
    }
    const users = await this.client.query(api.auth.listarUsuarios, {
      sessionToken: this.sessaoAtual.token,
    });
    this.usuariosAdmin = users as UsuarioAdmin[];
  }

  usuariosFiltrados(): UsuarioAdmin[] {
    const filtro = this.filtroUsuarios.trim().toLowerCase();
    if (!filtro) return this.usuariosAdmin;
    return this.usuariosAdmin.filter(
      (u) => u.nome.toLowerCase().includes(filtro) || u.email.toLowerCase().includes(filtro) || u.perfil.toLowerCase().includes(filtro),
    );
  }

  abrirConfirmacaoReset(user: UsuarioAdmin): void {
    this.confirmacaoModalAberto = true;
    this.confirmacaoAcao = "RESET_SENHA";
    this.confirmacaoUsuario = user;
    this.confirmacaoNovoAtivo = null;
    this.confirmacaoModalTitulo = "Confirmar reset de senha";
    this.confirmacaoModalMensagem = `Redefinir senha de ${user.email} para a temporaria Temp@123 e forcar troca no proximo login?`;
  }

  abrirConfirmacaoAtivo(user: UsuarioAdmin, novoAtivo: boolean): void {
    this.confirmacaoModalAberto = true;
    this.confirmacaoAcao = "ALTERAR_ATIVO";
    this.confirmacaoUsuario = user;
    this.confirmacaoNovoAtivo = novoAtivo;
    this.confirmacaoModalTitulo = novoAtivo ? "Confirmar ativacao" : "Confirmar inativacao";
    this.confirmacaoModalMensagem = novoAtivo
      ? `Ativar usuario ${user.email}?`
      : `Inativar usuario ${user.email}? Ele perdera acesso imediato.`;
  }

  fecharConfirmacao(): void {
    this.confirmacaoModalAberto = false;
    this.confirmacaoAcao = null;
    this.confirmacaoUsuario = null;
    this.confirmacaoNovoAtivo = null;
  }

  async confirmarAcaoModal(): Promise<void> {
    if (!this.confirmacaoUsuario || !this.confirmacaoAcao) {
      this.fecharConfirmacao();
      return;
    }

    try {
      if (this.confirmacaoAcao === "RESET_SENHA") {
        await this.redefinirSenhaUsuarioAdmin(this.confirmacaoUsuario);
      }

      if (this.confirmacaoAcao === "ALTERAR_ATIVO" && this.confirmacaoNovoAtivo !== null) {
        await this.client.mutation(api.auth.atualizarUsuario, {
          sessionToken: this.tokenSessao(),
          usuarioId: this.confirmacaoUsuario.id as any,
          nome: this.confirmacaoUsuario.nome,
          perfil: this.confirmacaoUsuario.perfil,
          ativo: this.confirmacaoNovoAtivo,
        });
        await this.carregarUsuariosAdmin();
      }
    } catch (e) {
      if (!(await this.tratarErroAutenticacao(e))) {
        this.authMensagem = e instanceof Error ? e.message : String(e);
      }
    } finally {
      this.fecharConfirmacao();
    }
  }

  async alterarSenhaObrigatoria(): Promise<void> {
    this.authMensagem = "";
    if (!this.sessaoAtual?.token) return;
    if (this.novaSenhaTroca !== this.confirmarSenhaTroca) {
      this.authMensagem = "A confirmacao da nova senha nao confere.";
      return;
    }

    try {
      await this.client.mutation(api.auth.alterarSenha, {
        sessionToken: this.sessaoAtual.token,
        senhaAtual: this.senhaAtualTroca,
        novaSenha: this.novaSenhaTroca,
      });

      const sessao = await this.client.query(api.auth.me, { sessionToken: this.sessaoAtual.token });
      this.sessaoAtual = sessao as SessaoAtual;
      this.senhaAtualTroca = "";
      this.novaSenhaTroca = "";
      this.confirmarSenhaTroca = "";
      this.authMensagem = "Senha alterada com sucesso.";
      await this.recarregar();
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.authMensagem = this.mensagemErroAuthAmigavel(e, "TROCA_SENHA");
    }
  }

  salvarPerfilOperacao(): void {
    localStorage.setItem("labre_operador", this.operador);
    localStorage.setItem("labre_perfil", this.perfil);
  }

  private fileToBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = String(reader.result ?? "");
        const base64 = result.includes(",") ? result.split(",")[1] : "";
        if (!base64) {
          reject(new Error("Falha ao converter arquivo para base64."));
          return;
        }
        resolve(base64);
      };
      reader.onerror = () => reject(new Error("Falha ao ler arquivo selecionado."));
      reader.readAsDataURL(file);
    });
  }

  private validarTamanhoArquivo(file: File): boolean {
    const maxBytes = 25 * 1024 * 1024;
    if (file.size > maxBytes) {
      this.uploadMensagem = `Arquivo ${file.name} excede 25MB.`;
      return false;
    }
    return true;
  }

  onArquivoSelecionado(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.arquivoSelecionado = input.files?.[0] ?? null;
    if (this.arquivoSelecionado && !this.validarTamanhoArquivo(this.arquivoSelecionado)) {
      this.arquivoSelecionado = null;
      return;
    }
    this.uploadMensagem = this.arquivoSelecionado ? `Arquivo selecionado: ${this.arquivoSelecionado.name}` : "";
  }

  onArquivoPavSelecionado(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.arquivoPavSelecionado = input.files?.[0] ?? null;
    if (this.arquivoPavSelecionado && !this.validarTamanhoArquivo(this.arquivoPavSelecionado)) {
      this.arquivoPavSelecionado = null;
    }
  }

  onArquivoNaoPavSelecionado(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.arquivoNaoPavSelecionado = input.files?.[0] ?? null;
    if (this.arquivoNaoPavSelecionado && !this.validarTamanhoArquivo(this.arquivoNaoPavSelecionado)) {
      this.arquivoNaoPavSelecionado = null;
    }
  }

  async importarArquivoSelecionado(): Promise<void> {
    if (!this.arquivoSelecionado) {
      this.uploadMensagem = "Selecione um arquivo antes de importar.";
      return;
    }

    this.uploadEmAndamento = true;
    this.uploadMensagem = "Enviando arquivo e processando no servidor...";

    try {
      const arquivoBase64 = await this.fileToBase64(this.arquivoSelecionado);

      const resultado = await this.client.mutation(api.trechos.iniciarImportacaoArquivoAsync, {
        sessionToken: this.tokenSessao(),
        tipoFonte: this.uploadTipoFonte,
        regiao: this.regiao,
        ano: this.ano,
        mes: this.mes,
        arquivoOrigem: this.arquivoSelecionado.name,
        arquivoBase64,
        limparAntes: this.uploadLimparAntes,
        dryRun: false,
        operador: this.operador || undefined,
        perfil: this.perfil,
      });

      this.uploadImportacaoId = String(resultado.importacaoId ?? "");
      this.uploadMensagem = `Importacao enfileirada (ID: ${this.uploadImportacaoId}). Aguardando processamento...`;

      if (this.uploadProcessarComplementar) {
        await this.client.action(api.workbook.importarWorkbookComplementar, {
          sessionToken: this.tokenSessao(),
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
          tipoFonte: this.uploadTipoFonte,
          arquivoOrigem: this.arquivoSelecionado.name,
          arquivoBase64,
          limparAntes: true,
          operador: this.operador || undefined,
          perfil: this.perfil,
        });
      }

      await this.acompanharImportacoes([this.uploadImportacaoId]);
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.uploadMensagem = e instanceof Error ? `Falha na importacao: ${e.message}` : String(e);
    } finally {
      this.uploadEmAndamento = false;
    }
  }

  async importarLoteSelecionado(): Promise<void> {
    if (!this.arquivoPavSelecionado || !this.arquivoNaoPavSelecionado) {
      this.uploadMensagem = "Selecione os dois arquivos (PAV e NAO_PAV).";
      return;
    }

    this.uploadEmAndamento = true;
    this.uploadMensagem = "Enviando lote PAV + NAO_PAV para processamento...";

    try {
      const [pavArquivoBase64, naoPavArquivoBase64] = await Promise.all([
        this.fileToBase64(this.arquivoPavSelecionado),
        this.fileToBase64(this.arquivoNaoPavSelecionado),
      ]);

      const resultado = await this.client.mutation(api.trechos.iniciarImportacaoLoteAsync, {
        sessionToken: this.tokenSessao(),
        regiao: this.regiao,
        ano: this.ano,
        mes: this.mes,
        pavArquivoOrigem: this.arquivoPavSelecionado.name,
        pavArquivoBase64,
        naoPavArquivoOrigem: this.arquivoNaoPavSelecionado.name,
        naoPavArquivoBase64,
        limparAntes: this.uploadLimparAntes,
        dryRun: false,
        operador: this.operador || undefined,
        perfil: this.perfil,
      });

      const ids = [String(resultado.importacoes.PAV), String(resultado.importacoes.NAO_PAV)];
      this.uploadMensagem = `Lote enfileirado. IDs: ${ids.join(", ")}`;

      if (this.uploadProcessarComplementar) {
        await this.client.action(api.workbook.importarWorkbookComplementar, {
          sessionToken: this.tokenSessao(),
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
          tipoFonte: "PAV",
          arquivoOrigem: this.arquivoPavSelecionado.name,
          arquivoBase64: pavArquivoBase64,
          limparAntes: true,
          operador: this.operador || undefined,
          perfil: this.perfil,
        });
        await this.client.action(api.workbook.importarWorkbookComplementar, {
          sessionToken: this.tokenSessao(),
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
          tipoFonte: "NAO_PAV",
          arquivoOrigem: this.arquivoNaoPavSelecionado.name,
          arquivoBase64: naoPavArquivoBase64,
          limparAntes: true,
          operador: this.operador || undefined,
          perfil: this.perfil,
        });
      }

      await this.acompanharImportacoes(ids);
      this.arquivoPavSelecionado = null;
      this.arquivoNaoPavSelecionado = null;
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.uploadMensagem = e instanceof Error ? `Falha no lote: ${e.message}` : String(e);
    } finally {
      this.uploadEmAndamento = false;
    }
  }

  private async acompanharImportacoes(importacaoIds: string[]): Promise<void> {
    const espera = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

    for (let tentativa = 0; tentativa < 120; tentativa += 1) {
      const statusList = (await Promise.all(
        importacaoIds.map((id) =>
          this.client.query(api.trechos.obterStatusImportacao, {
            importacaoId: id as any,
          }),
        ),
      )) as ImportacaoStatus[];

      const pendentes = statusList.filter((s: ImportacaoStatus) => s.status === "PROCESSANDO");
      if (pendentes.length > 0) {
        this.uploadMensagem = `Processando ${pendentes.length}/${statusList.length} importacao(oes)...`;
        await espera(1500);
        continue;
      }

      const totalGravados = statusList.reduce((acc: number, s: ImportacaoStatus) => acc + s.gravados, 0);
      const totalIgnoradas = statusList.reduce((acc: number, s: ImportacaoStatus) => acc + s.linhasIgnoradas, 0);
      const totalErros = statusList.reduce((acc: number, s: ImportacaoStatus) => acc + s.linhasComErro, 0);
      const erros = statusList.filter((s: ImportacaoStatus) => s.status === "ERRO");

      if (erros.length > 0) {
        this.uploadMensagem = `Concluido com falha em ${erros.length} importacao(oes). Consulte o historico.`;
      } else {
        this.uploadMensagem = `Importacao concluida: ${totalGravados} gravados, ${totalIgnoradas} ignoradas, ${totalErros} erros.`;
      }

      this.arquivoSelecionado = null;
      await this.recarregar();
      return;
    }

    this.uploadMensagem = "Processamento ainda em andamento. Atualize em alguns segundos.";
  }

  async recarregar(): Promise<void> {
    this.loading = true;
    this.erro = "";
    try {
      const isAdmin = this.sessaoAtual?.usuario?.perfil === "ADMIN";
      const [graficos, inconsistencias, auditoria, saude, evolucao] = await Promise.all([
        this.client.query(api.trechos.obterGraficosCompetencia, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        this.client.query(api.trechos.obterInconsistenciasImportacao, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        isAdmin
          ? this.client.query(api.trechos.listarAuditoriaRecente, {
              sessionToken: this.tokenSessao(),
              limite: 30,
            })
          : Promise.resolve([]),
        isAdmin
          ? this.client.query(api.trechos.obterSaudeOperacional, {
              sessionToken: this.tokenSessao(),
            })
          : Promise.resolve(null),
        this.client.query(api.trechos.obterEvolucaoManutencao, {
          regiao: this.regiao,
          ano: this.ano,
          trecho: this.evolucaoTrechoSelecionado || undefined,
        }),
      ]);

      this.totalTrechos = graficos.kpis.totalTrechos;
      this.totalKm = graficos.kpis.totalKm;
      this.programados = graficos.kpis.programadosNoMes;
      this.naoProgramados = graficos.kpis.naoProgramadosNoMes;

      this.porTipoFonte = graficos.series.porTipoFonte;
      this.topSrePorKm = graficos.series.topSrePorKm;

      this.importacoes = inconsistencias.importacoes as ImportacaoResumo[];
      this.porCodigoInconsistencia = inconsistencias.porCodigo as CodigoInconsistenciaResumo[];
      this.resumoInconsistencias = inconsistencias.resumo;
      this.auditoriaRecente = (auditoria as AuditoriaEvento[]) ?? [];
      this.saudeOperacional = (saude as SaudeOperacional | null) ?? null;
      const evolucaoPayload = evolucao as {
        trechoSelecionado: string;
        trechosDisponiveis: string[];
        mensal: EvolucaoManutencaoItem[];
      };
      this.evolucaoTrechosDisponiveis = evolucaoPayload.trechosDisponiveis;
      this.evolucaoTrechoSelecionado = evolucaoPayload.trechoSelecionado;
      this.evolucaoMensal = evolucaoPayload.mensal;
      await this.carregarUsuariosAdmin();
      await this.recarregarGraficosWorkbook();
    } catch (e) {
      if (await this.tratarErroAutenticacao(e)) return;
      this.erro = e instanceof Error ? e.message : String(e);
    } finally {
      this.loading = false;
    }
  }

  async recarregarEvolucaoManutencao(): Promise<void> {
    if (!this.sessaoAtual?.token) return;
    try {
      const evolucao = (await this.client.query(api.trechos.obterEvolucaoManutencao, {
        regiao: this.regiao,
        ano: this.ano,
        trecho: this.evolucaoTrechoSelecionado || undefined,
      })) as {
        trechoSelecionado: string;
        trechosDisponiveis: string[];
        mensal: EvolucaoManutencaoItem[];
      };

      this.evolucaoTrechosDisponiveis = evolucao.trechosDisponiveis;
      this.evolucaoTrechoSelecionado = evolucao.trechoSelecionado;
      this.evolucaoMensal = evolucao.mensal;
    } catch (e) {
      if (!(await this.tratarErroAutenticacao(e))) {
        this.erro = e instanceof Error ? e.message : String(e);
      }
    }
  }

  maxKm(): number {
    const max = this.topSrePorKm.reduce((acc, item) => Math.max(acc, item.km), 0);
    return max || 1;
  }

  barraKm(valor: number): string {
    return `${Math.max(4, (valor / this.maxKm()) * 100)}%`;
  }

  maxInconsistenciasPorCodigo(): number {
    const max = this.porCodigoInconsistencia.reduce((acc, item) => Math.max(acc, item.total), 0);
    return max || 1;
  }

  barraInconsistencia(total: number): string {
    return `${Math.max(4, (total / this.maxInconsistenciasPorCodigo()) * 100)}%`;
  }

  importacoesOrdenadas(): ImportacaoResumo[] {
    return [...this.importacoes].sort((a, b) => {
      const aTempo = a.finalizadoEm ?? a.iniciadoEm ?? 0;
      const bTempo = b.finalizadoEm ?? b.iniciadoEm ?? 0;
      return bTempo - aTempo;
    });
  }

  ultimasImportacoesPorTipo(): ImportacaoResumo[] {
    const latestByTipo = new Map<string, ImportacaoResumo>();

    for (const item of this.importacoesOrdenadas()) {
      if (!latestByTipo.has(item.tipoFonte)) {
        latestByTipo.set(item.tipoFonte, item);
      }
    }

    return Array.from(latestByTipo.values());
  }

  importacoesVisiveis(): ImportacaoResumo[] {
    return this.mostrarHistoricoCompleto ? this.importacoesOrdenadas() : this.ultimasImportacoesPorTipo();
  }

  totalIgnoradasUltimasImportacoes(): number {
    return this.ultimasImportacoesPorTipo().reduce((acc, item) => acc + item.linhasIgnoradas, 0);
  }

  totalRecebidasUltimasImportacoes(): number {
    return this.ultimasImportacoesPorTipo().reduce(
      (acc, item) => acc + item.linhasValidas + item.linhasIgnoradas + item.linhasComErro,
      0,
    );
  }

  alertaIgnoradasElevadas(): boolean {
    const ignoradas = this.totalIgnoradasUltimasImportacoes();
    const recebidas = this.totalRecebidasUltimasImportacoes();
    if (recebidas === 0) return false;

    const percentualIgnoradas = ignoradas / recebidas;
    return ignoradas >= 20 && percentualIgnoradas >= 0.25;
  }

  textoAlertaIgnoradas(): string {
    const ignoradas = this.totalIgnoradasUltimasImportacoes();
    const recebidas = this.totalRecebidasUltimasImportacoes();
    const percentual = recebidas === 0 ? 0 : Math.round((ignoradas / recebidas) * 100);

    return `${ignoradas} linhas ignoradas (${percentual}% da carga recente). Verifique se a planilha possui linhas de outras regioes.`;
  }

  private percent(part: number, total: number): number {
    if (total <= 0) return 0;
    return Number(((part / total) * 100).toFixed(1));
  }

  dadosPizzaTipoFonte(): Array<{ label: string; valor: number; percentual: number; color: string }> {
    const total = this.porTipoFonte.reduce((acc, item) => acc + item.totalTrechos, 0);
    const palette = ["#d95f02", "#1b9e77", "#457b9d", "#e9c46a"];
    return this.porTipoFonte.map((item, idx) => ({
      label: item.tipoFonte,
      valor: item.totalTrechos,
      percentual: this.percent(item.totalTrechos, total),
      color: palette[idx % palette.length],
    }));
  }

  dadosPizzaProgramacao(): Array<{ label: string; valor: number; percentual: number; color: string }> {
    const total = this.programados + this.naoProgramados;
    return [
      {
        label: "Programados",
        valor: this.programados,
        percentual: this.percent(this.programados, total),
        color: "#2a9d8f",
      },
      {
        label: "Nao programados",
        valor: this.naoProgramados,
        percentual: this.percent(this.naoProgramados, total),
        color: "#e76f51",
      },
    ];
  }

  estiloPizzaConica(data: Array<{ percentual: number; color: string }>): string {
    let angle = 0;
    const slices = data
      .map((item) => {
        const start = angle;
        const inc = (item.percentual / 100) * 360;
        angle += inc;
        return `${item.color} ${start}deg ${angle}deg`;
      })
      .join(", ");
    return `conic-gradient(${slices || "#e2e8f0 0deg 360deg"})`;
  }

  rotulosPizza(
    data: Array<{ valor: number; percentual: number }>,
  ): Array<{ texto: string; left: string; top: string }> {
    const cx = 50;
    const cy = 50;
    const pieRadius = 47;
    let angulo = -90;

    const slices = data
      .filter((item) => item.percentual > 0)
      .map((item) => {
        const span = (item.percentual / 100) * 360;
        const start = angulo;
        const meio = start + span / 2;
        angulo += span;
        return {
          valor: item.valor,
          percentual: item.percentual,
          start,
          span,
          meio,
          texto: `${this.formatarNumero(item.valor, 1)}\n${this.formatarNumero(item.percentual, 1)}%`,
        };
      })
      .sort((a, b) => b.percentual - a.percentual);

    const placed: Array<{ x: number; y: number; w: number; h: number }> = [];
    const externos: Array<{ texto: string; x: number; y: number; w: number; h: number; side: "left" | "right" }> = [];
    const output: Array<{ texto: string; left: string; top: string }> = [];

    const collide = (x: number, y: number, w: number, h: number): boolean =>
      placed.some((p) => Math.abs(p.x - x) < (p.w + w) / 2 && Math.abs(p.y - y) < (p.h + h) / 2);

    for (const slice of slices) {
      const linha1 = this.formatarNumero(slice.valor, 1);
      const linha2 = `${this.formatarNumero(slice.percentual, 1)}%`;
      const maxChars = Math.max(linha1.length, linha2.length);
      const w = Math.min(24, Math.max(15, 8 + maxChars * 1.45));
      const h = 24;
      const outsidePreferred = slice.percentual < 12;

      const radiusOptions = [pieRadius * 0.6, pieRadius * 0.5, pieRadius * 0.68, pieRadius * 0.4];
      const offsetOptions = [0, -slice.span * 0.22, slice.span * 0.22, -slice.span * 0.35, slice.span * 0.35];

      let chosen: { x: number; y: number } | null = null;
      if (!outsidePreferred) {
        for (const r of radiusOptions) {
          for (const delta of offsetOptions) {
            const ang = ((slice.meio + delta) * Math.PI) / 180;
            const x = cx + r * Math.cos(ang);
            const y = cy + r * Math.sin(ang);
            const dist = Math.hypot(x - cx, y - cy);
            const maxDist = pieRadius - Math.max(w, h) * 0.25;
            if (dist > maxDist) continue;
            if (collide(x, y, w, h)) continue;
            chosen = { x, y };
            break;
          }
          if (chosen) break;
        }
      }

      if (chosen) {
        placed.push({ x: chosen.x, y: chosen.y, w, h });
        output.push({
          texto: slice.texto,
          left: `${chosen.x}%`,
          top: `${chosen.y}%`,
        });
        continue;
      }

      const rad = (slice.meio * Math.PI) / 180;
      externos.push({
        texto: slice.texto,
        x: cx + pieRadius * 1.1 * Math.cos(rad),
        y: cy + pieRadius * 1.1 * Math.sin(rad),
        w,
        h,
        side: Math.cos(rad) >= 0 ? "right" : "left",
      });
    }

    const ajustarExternos = (side: "left" | "right") => {
      const itens = externos.filter((e) => e.side === side).sort((a, b) => a.y - b.y);
      const minY = 8;
      const maxY = 92;

      for (let i = 0; i < itens.length; i += 1) {
        if (i > 0) {
          const minGap = Math.max(itens[i].h, itens[i - 1].h) + 2;
          if (itens[i].y - itens[i - 1].y < minGap) itens[i].y = itens[i - 1].y + minGap;
        }
      }
      for (let i = itens.length - 1; i >= 0; i -= 1) {
        const top = minY + itens[i].h / 2;
        const bottom = maxY - itens[i].h / 2;
        if (itens[i].y > bottom) itens[i].y = bottom;
        if (itens[i].y < top) itens[i].y = top;
        if (i > 0) {
          const minGap = Math.max(itens[i].h, itens[i - 1].h) + 2;
          if (itens[i].y - itens[i - 1].y < minGap) itens[i - 1].y = itens[i].y - minGap;
        }
      }

      for (const item of itens) {
        const top = minY + item.h / 2;
        const bottom = maxY - item.h / 2;
        if (item.y < top) item.y = top;
        if (item.y > bottom) item.y = bottom;
        if (side === "right") {
          item.x = Math.max(62, Math.min(78, item.x));
        } else {
          item.x = Math.max(22, Math.min(38, item.x));
        }
      }
    };

    ajustarExternos("left");
    ajustarExternos("right");

    for (const item of externos) {
      output.push({ texto: item.texto, left: `${item.x}%`, top: `${item.y}%` });
    }

    return output;
  }

  rotulosPizzaTipoFonte(): Array<{ texto: string; left: string; top: string }> {
    return this.rotulosPizza(this.dadosPizzaTipoFonte());
  }

  rotulosPizzaProgramacao(): Array<{ texto: string; left: string; top: string }> {
    return this.rotulosPizza(this.dadosPizzaProgramacao());
  }

  rotulosPizzaWorkbook(grafico: WorkbookGrafico): Array<{ texto: string; left: string; top: string }> {
    return this.rotulosPizza(grafico.series.map((s) => ({ valor: s.valor, percentual: s.percentual })));
  }

  private paletteWorkbook = ["#d95f02", "#1b9e77", "#457b9d", "#e9c46a", "#e76f51", "#6d597a", "#264653"];

  workbookTrechosDisponiveis(): string[] {
    return Array.from(new Set(this.workbookGraficos.map((g) => g.trecho?.trim()).filter((v) => !!v) as string[])).sort(
      (a, b) => a.localeCompare(b, "pt-BR"),
    );
  }

  private isTrechoNaoInformado(value?: string): boolean {
    if (!value) return true;
    return this.normalizarTextoRelatorio(value) === "trecho nao informado";
  }

  private resumoAnalisePorTipo(tipo: "PAV" | "NAO_PAV"): string {
    return tipo === "PAV"
      ? `Resumo de analise da Regiao ${String(this.regiao).padStart(2, "0")} - Rodovias Pavimentadas`
      : `Resumo de analise da Regiao ${String(this.regiao).padStart(2, "0")} - Rodovias Nao Pavimentadas`;
  }

  rotuloTrechoWorkbook(trecho?: string, tipoFonte?: "PAV" | "NAO_PAV"): string {
    const tipo = tipoFonte ?? this.workbookTipoFonte;
    if (this.isTrechoNaoInformado(trecho)) {
      return this.resumoAnalisePorTipo(tipo);
    }
    return trecho?.trim() || "-";
  }

  graficosWorkbookFiltrados(): WorkbookGrafico[] {
    if (!this.workbookTrechoSelecionado) return this.workbookGraficos;
    return this.workbookGraficos.filter((g) => (g.trecho ?? "") === this.workbookTrechoSelecionado);
  }

  estiloPizzaWorkbook(grafico: WorkbookGrafico): string {
    const data = grafico.series.map((s, i) => ({ percentual: s.percentual, color: this.paletteWorkbook[i % this.paletteWorkbook.length] }));
    return this.estiloPizzaConica(data);
  }

  corSerieWorkbook(index: number): string {
    return this.paletteWorkbook[index % this.paletteWorkbook.length];
  }

  maxEvolucaoKm(): number {
    const max = this.evolucaoMensal.reduce((acc, item) => Math.max(acc, item.geralKm, item.trechoKm), 0);
    return max || 1;
  }

  barraEvolucao(valor: number): string {
    return `${Math.max(2, (valor / this.maxEvolucaoKm()) * 100)}%`;
  }

  async recarregarGraficosWorkbook(): Promise<void> {
    if (!this.sessaoAtual?.token) return;

    const result = await this.client.query(api.workbook.listarGraficosWorkbook, {
      sessionToken: this.tokenSessao(),
      regiao: this.regiao,
      ano: this.ano,
      mes: this.mes,
      tipoFonte: this.workbookTipoFonte,
    });

    this.workbookGraficos = result as WorkbookGrafico[];
    const trechos = this.workbookTrechosDisponiveis();
    if (trechos.length === 0) {
      this.workbookTrechoSelecionado = "";
    } else if (!trechos.includes(this.workbookTrechoSelecionado)) {
      this.workbookTrechoSelecionado = trechos[0];
    }
  }

  formatarDataImportacao(item: ImportacaoResumo): string {
    const ts = item.finalizadoEm ?? item.iniciadoEm;
    if (!ts) return "Sem horario";

    return new Intl.DateTimeFormat("pt-BR", {
      dateStyle: "short",
      timeStyle: "short",
    }).format(ts);
  }

  nomeArquivoMarkdown(): string {
    return `relatorio_regiao_${this.regiao}_${this.competencia()}.md`;
  }

  async onEstruturaRelatorioSelecionada(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0] ?? null;
    if (!file) return;

    try {
      const raw = await file.text();
      const parsed = JSON.parse(raw) as { headings?: unknown };
      if (!Array.isArray(parsed.headings)) {
        throw new Error("JSON invalido: propriedade 'headings' ausente.");
      }

      const headings = parsed.headings
        .map((item) => this.parseTemplateHeading(item))
        .filter((item): item is RelatorioTemplateHeading => item !== null);

      if (headings.length === 0) {
        throw new Error("Nenhum heading valido encontrado no JSON informado.");
      }

      this.templateEstruturaHeadings = this.deduplicarHeadingsTemplate(headings);
      this.templateEstruturaNome = file.name;
      this.templateEstruturaMensagem = `Estrutura carregada: ${this.templateEstruturaHeadings.length} secoes unicas (${file.name}).`;
    } catch (e) {
      this.templateEstruturaHeadings = [];
      this.templateEstruturaNome = "";
      this.templateEstruturaMensagem =
        e instanceof Error
          ? `Falha ao carregar estrutura: ${e.message}`
          : `Falha ao carregar estrutura: ${String(e)}`;
    } finally {
      input.value = "";
    }
  }

  onTemplateDocxSelecionado(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0] ?? null;
    if (!file) return;

    if (!file.name.toLowerCase().endsWith(".docx")) {
      this.templateDocxBase = null;
      this.templateDocxMensagem = "Arquivo invalido: selecione um .docx de relatorio base.";
      input.value = "";
      return;
    }

    this.templateDocxBase = file;
    this.templateDocxMensagem = `Template base carregado: ${file.name}. O DOCX regional vai preservar capa/cabecalho e atualizar apenas campos dinamicos e graficos.`;
    input.value = "";
  }

  private parseTemplateHeading(item: unknown): RelatorioTemplateHeading | null {
    if (!item || typeof item !== "object") return null;
    const value = item as Record<string, unknown>;
    const texto = typeof value.texto === "string" ? value.texto.trim() : "";
    const nivel = typeof value.nivel === "number" ? value.nivel : Number(value.nivel);
    const index = typeof value.index === "number" ? value.index : Number(value.index);
    const origemRaw = typeof value.origem === "string" ? value.origem : "numeracao";
    const origem = origemRaw === "estilo" ? "estilo" : "numeracao";
    if (!texto || !Number.isFinite(nivel) || !Number.isFinite(index)) return null;
    return {
      index,
      nivel: Math.max(1, Math.min(6, Math.trunc(nivel))),
      texto,
      origem,
    };
  }

  private deduplicarHeadingsTemplate(headings: RelatorioTemplateHeading[]): RelatorioTemplateHeading[] {
    const ordered = [...headings].sort((a, b) => a.index - b.index);
    const seen = new Set<string>();
    const result: RelatorioTemplateHeading[] = [];

    for (const heading of ordered) {
      const key = this.normalizarTextoRelatorio(heading.texto);
      if (!key || seen.has(key)) continue;
      seen.add(key);
      result.push(heading);
    }

    return result;
  }

  private normalizarTextoRelatorio(value: string): string {
    return value
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  }

  private headingTemplateOuPadrao(fallback: string, pistas: string[]): string {
    if (this.templateEstruturaHeadings.length === 0) return fallback;
    const pistasNorm = pistas.map((p) => this.normalizarTextoRelatorio(p));

    for (const heading of this.templateEstruturaHeadings) {
      const textoNorm = this.normalizarTextoRelatorio(heading.texto);
      if (pistasNorm.every((pista) => textoNorm.includes(pista))) {
        return heading.texto;
      }
    }

    return fallback;
  }

  private linhasMarkdownOperacionais(): string[] {
    return [
      "## Indicadores",
      `- Total de trechos: ${this.totalTrechos}`,
      `- Total de extensao (km): ${this.totalKm.toFixed(2)}`,
      `- Programados no mes: ${this.programados}`,
      `- Nao programados no mes: ${this.naoProgramados}`,
      "",
      "## Distribuicao por Tipo de Fonte",
      ...this.porTipoFonte.map(
        (item) => `- ${item.tipoFonte}: ${item.totalTrechos} trechos / ${item.totalKm.toFixed(2)} km`,
      ),
      "",
      "## Top SRE por KM",
      ...this.topSrePorKm.map((item) => `- ${item.sre}: ${item.km.toFixed(2)} km`),
      "",
      "## Inconsistencias de Importacao",
      `- Total de importacoes: ${this.resumoInconsistencias.totalImportacoes}`,
      `- Importacoes com erro: ${this.resumoInconsistencias.importacoesComErro}`,
      `- Total de erros: ${this.resumoInconsistencias.totalErros}`,
      ...this.porCodigoInconsistencia.map((item) => `- ${item.codigo}: ${item.total}`),
      "",
      "## Historico de Importacoes",
      ...this.importacoesOrdenadas().flatMap((item) => [
        `- ${this.formatarDataImportacao(item)} | ${item.tipoFonte} | ${item.status}`,
        `  - Arquivo: ${item.arquivoOrigem}`,
        `  - Validas: ${item.linhasValidas} | Ignoradas: ${item.linhasIgnoradas} | Erros: ${item.linhasComErro} | Gravados: ${item.gravados}`,
      ]),
      "",
    ];
  }

  private linhasResumoTipoFonte(tipoFonte: "PAV" | "NAO_PAV"): string[] {
    const fonte = this.porTipoFonte.find((f) => f.tipoFonte === tipoFonte);
    const importacoes = this.importacoes.filter((item) => item.tipoFonte === tipoFonte);
    const linhasValidas = importacoes.reduce((acc, item) => acc + item.linhasValidas, 0);
    const linhasIgnoradas = importacoes.reduce((acc, item) => acc + item.linhasIgnoradas, 0);
    const linhasComErro = importacoes.reduce((acc, item) => acc + item.linhasComErro, 0);

    return [
      `- Total de trechos: ${fonte?.totalTrechos ?? 0}`,
      `- Total de extensao (km): ${(fonte?.totalKm ?? 0).toFixed(2)}`,
      `- Importacoes no periodo: ${importacoes.length}`,
      `- Linhas validas: ${linhasValidas}`,
      `- Linhas ignoradas: ${linhasIgnoradas}`,
      `- Linhas com erro: ${linhasComErro}`,
    ];
  }

  private linhasTrechosTemplateMarkdown(tipo: "PAV" | "NAO_PAV", graficos: WorkbookGrafico[]): string[] {
    const gruposOrdenados = this.ordenarGruposPorTemplate(tipo, this.agruparGraficosPorTrecho(graficos));
    if (gruposOrdenados.length > 0) {
      return gruposOrdenados.map((grupo, idx) => {
        const titulo = grupo.tituloTemplate ?? this.rotuloTrechoWorkbook(grupo.trecho, tipo);
        return `- ${idx + 1}. ${titulo} | ${grupo.graficos.length} grafico(s)`;
      });
    }

    const trechosTemplate = this.ordemTrechosTemplate(tipo);
    if (trechosTemplate.length > 0) {
      return trechosTemplate.map((trecho, idx) => `- ${idx + 1}. ${trecho}`);
    }

    return ["- Sem dados de trecho para este tipo de fonte na competencia atual."];
  }

  private async carregarGraficosWorkbookMarkdown(): Promise<{ pav: WorkbookGrafico[]; naoPav: WorkbookGrafico[] }> {
    const [pavRaw, naoPavRaw] = await Promise.all([
      this.client.query(api.workbook.listarGraficosWorkbook, {
        sessionToken: this.tokenSessao(),
        regiao: this.regiao,
        ano: this.ano,
        mes: this.mes,
        tipoFonte: "PAV",
      }),
      this.client.query(api.workbook.listarGraficosWorkbook, {
        sessionToken: this.tokenSessao(),
        regiao: this.regiao,
        ano: this.ano,
        mes: this.mes,
        tipoFonte: "NAO_PAV",
      }),
    ]);

    return {
      pav: pavRaw as WorkbookGrafico[],
      naoPav: naoPavRaw as WorkbookGrafico[],
    };
  }

  private linhasParaHeadingTemplate(textoHeading: string): string[] {
    const normalizado = this.normalizarTextoRelatorio(textoHeading);
    if (normalizado.includes("apresentacao")) {
      return [
        `Este relatorio sintetiza os resultados da competencia ${this.competencia()} para a Regiao ${this.regiao}, com foco em produtividade da manutencao, qualidade de dados e consistencia operacional.`,
      ];
    }

    if (normalizado.includes("vistorias rotineiras no periodo")) {
      return [
        ...this.linhasResumoTipoFonte("PAV"),
        "",
        ...this.linhasResumoTipoFonte("NAO_PAV"),
      ];
    }

    if (normalizado.includes("rodovias pavimentadas") && !normalizado.includes("nao pavimentadas")) {
      return this.linhasResumoTipoFonte("PAV");
    }

    if (normalizado.includes("rodovias nao pavimentadas")) {
      return this.linhasResumoTipoFonte("NAO_PAV");
    }

    if (normalizado.includes("resumo de analise do geral")) {
      return [
        `- Total de trechos: ${this.totalTrechos}`,
        `- Total de extensao (km): ${this.totalKm.toFixed(2)}`,
        `- Programados no mes: ${this.programados}`,
        `- Nao programados no mes: ${this.naoProgramados}`,
        ...this.topSrePorKm.slice(0, 8).map((item) => `- Top SRE: ${item.sre} (${item.km.toFixed(2)} km)`),
      ];
    }

    if (normalizado.includes("termo de encerramento")) {
      return [
        `- Competencia consolidada: ${this.competencia()}`,
        `- Data de emissao automatica: ${new Date().toISOString()}`,
        "- Resultado tecnico: dados processados e relatorio emitido sem erro fatal no frontend.",
      ];
    }

    return [];
  }

  private gerarMarkdownCompetenciaComTemplate(): string {
    const linhas: string[] = [
      `# Relatorio Tecnico - Regiao ${this.regiao}`,
      "",
      `- Competencia: ${this.competencia()}`,
      `- Gerado em: ${new Date().toISOString()}`,
      `- Estrutura base: ${this.templateEstruturaNome}`,
      "",
    ];

    let blocosDinamicosInseridos = 0;
    for (const heading of this.templateEstruturaHeadings) {
      const nivel = Math.max(2, Math.min(6, heading.nivel + 1));
      linhas.push(`${"#".repeat(nivel)} ${heading.texto}`);
      const bloco = this.linhasParaHeadingTemplate(heading.texto);
      if (bloco.length > 0) {
        linhas.push("");
        linhas.push(...bloco);
        blocosDinamicosInseridos += 1;
      }
      linhas.push("");
    }

    if (blocosDinamicosInseridos === 0) {
      linhas.push("## Dados Operacionais Complementares");
      linhas.push("");
      linhas.push(...this.linhasMarkdownOperacionais());
    } else {
      linhas.push("## Anexo Operacional");
      linhas.push("");
      linhas.push(...this.linhasMarkdownOperacionais());
    }

    return linhas.join("\n");
  }

  gerarMarkdownCompetencia(): string {
    if (this.templateEstruturaHeadings.length > 0) {
      return this.gerarMarkdownCompetenciaComTemplate();
    }

    const linhas = [
      `# Relatorio Tecnico - Regiao ${this.regiao}`,
      "",
      `- Competencia: ${this.competencia()}`,
      `- Gerado em: ${new Date().toISOString()}`,
      "",
      ...this.linhasMarkdownOperacionais(),
    ];

    return linhas.join("\n");
  }

  private gerarAnexoMarkdownTrechosTemplate(graficos: { pav: WorkbookGrafico[]; naoPav: WorkbookGrafico[] }): string {
    const linhas = [
      "## Anexo de Trechos (ordem do template)",
      "",
      "### Rodovias Pavimentadas",
      ...this.linhasTrechosTemplateMarkdown("PAV", graficos.pav),
      "",
      "### Rodovias Nao Pavimentadas",
      ...this.linhasTrechosTemplateMarkdown("NAO_PAV", graficos.naoPav),
      "",
    ];

    return linhas.join("\n");
  }

  async baixarMarkdownCompetencia(): Promise<void> {
    try {
      let markdown = this.gerarMarkdownCompetencia();
      if (this.templateEstruturaHeadings.length > 0) {
        const graficos = await this.carregarGraficosWorkbookMarkdown();
        markdown = `${markdown}\n${this.gerarAnexoMarkdownTrechosTemplate(graficos)}`;
      }

      const blob = new Blob([markdown], { type: "text/markdown;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = this.nomeArquivoMarkdown();
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      this.erro = e instanceof Error ? e.message : String(e);
    }
  }

  private async baixarDocx(nomeArquivo: string, linhas: string[]): Promise<void> {
    const children: Paragraph[] = [];
    for (const linha of linhas) {
      if (linha.startsWith("# ")) {
        children.push(new Paragraph({ text: linha.replace(/^#\s+/, ""), heading: HeadingLevel.TITLE }));
      } else if (linha.startsWith("## ")) {
        children.push(new Paragraph({ text: linha.replace(/^##\s+/, ""), heading: HeadingLevel.HEADING_1 }));
      } else if (linha.startsWith("- ")) {
        children.push(new Paragraph({ text: linha.slice(2), bullet: { level: 0 } }));
      } else if (linha.trim() === "") {
        children.push(new Paragraph({ text: "" }));
      } else {
        children.push(new Paragraph({ text: linha }));
      }
    }

    const doc = new Document({ sections: [{ children }] });
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = nomeArquivo;
    a.click();
    URL.revokeObjectURL(url);
  }

  private mesLabel(value: number): string {
    return this.meses.find((m) => m.value === value)?.label ?? String(value);
  }

  private formatarNumero(value: number, casas = 2): string {
    return new Intl.NumberFormat("pt-BR", {
      minimumFractionDigits: casas,
      maximumFractionDigits: casas,
    }).format(value);
  }

  private formatarDataHora(ts?: number): string {
    if (!ts) return "-";
    return new Intl.DateTimeFormat("pt-BR", {
      dateStyle: "short",
      timeStyle: "short",
    }).format(ts);
  }

  private dataUrlParaBytes(dataUrl: string): Uint8Array {
    const base64 = dataUrl.split(",")[1] ?? "";
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }

  private async bytesParaImagem(bytes: Uint8Array): Promise<HTMLImageElement> {
    const arrayBuffer = bytes.slice().buffer as ArrayBuffer;
    const blob = new Blob([arrayBuffer], { type: "image/png" });
    const url = URL.createObjectURL(blob);
    try {
      return await new Promise<HTMLImageElement>((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => reject(new Error("Falha ao carregar imagem temporaria."));
        img.src = url;
      });
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  private async renderizarPizzaWorkbook(
    titulo: string,
    series: WorkbookSerie[],
    opts?: { width?: number; height?: number },
  ): Promise<Uint8Array> {
    const width = opts?.width ?? 620;
    const height = opts?.height ?? 395;
    const canvas = document.createElement("canvas");
    const dpr = 2;
    canvas.width = width * dpr;
    canvas.height = height * dpr;
    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Falha ao inicializar contexto grafico do relatorio.");
    ctx.scale(dpr, dpr);

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);

    const titleFont = 15;
    const legendFont = 14;
    const labelFont = Math.max(12, Math.round(height * 0.054));
    const legendW = Math.max(130, Math.round(width * 0.3));
    const titleMaxW = width - legendW - 18;

    const drawWrappedTitle = (text: string, x: number, y: number, maxWidth: number, maxLines: number) => {
      const words = text.split(/\s+/).filter(Boolean);
      const lines: string[] = [];
      let current = "";
      for (const word of words) {
        const test = current ? `${current} ${word}` : word;
        if (ctx.measureText(test).width <= maxWidth || !current) {
          current = test;
        } else {
          lines.push(current);
          current = word;
          if (lines.length === maxLines - 1) break;
        }
      }
      if (current && lines.length < maxLines) lines.push(current);
      if (words.length > 0 && lines.length > 0 && lines.length === maxLines) {
        const full = words.join(" ");
        const shown = lines.join(" ");
        if (full.length > shown.length && !lines[maxLines - 1].endsWith("...")) {
          lines[maxLines - 1] = `${lines[maxLines - 1].replace(/\.*$/, "")}...`;
        }
      }
      for (let i = 0; i < lines.length; i += 1) {
        ctx.fillText(lines[i], x, y + i * Math.round(titleFont * 1.1));
      }
    };

    ctx.fillStyle = "#1f2933";
    ctx.font = `bold ${titleFont}px DM Sans, Segoe UI, sans-serif`;
    drawWrappedTitle(titulo, 14, Math.round(titleFont + 8), titleMaxW, 3);

    const total = series.reduce((acc, s) => acc + (s.valor ?? 0), 0);
    const plotLeft = 16;
    const plotTop = Math.round(titleFont * 2.2);
    const plotRight = width - legendW - 12;
    const legendX = width - legendW + 8;
    const plotBottom = height - 14;
    const plotW = Math.max(180, plotRight - plotLeft);
    const plotH = Math.max(160, plotBottom - plotTop);
    const cx = Math.round(plotLeft + plotW * 0.46);
    const cy = Math.round(plotTop + plotH * 0.56);
    const radius = Math.round(Math.min(plotW * 0.34, plotH * 0.42));
    let current = -Math.PI / 2;

    const labels: Array<{ linhas: [string, string]; x: number; y: number; w: number; h: number }> = [];
    const labelsExternos: Array<{
      linhas: [string, string];
      x: number;
      y: number;
      w: number;
      h: number;
      side: "left" | "right";
      anchorX: number;
      anchorY: number;
    }> = [];
    const collide = (x: number, y: number, w: number, h: number): boolean =>
      labels.some((l) => Math.abs(l.x - x) < (l.w + w) / 2 && Math.abs(l.y - y) < (l.h + h) / 2);
    ctx.font = `bold ${labelFont}px DM Sans, Segoe UI, sans-serif`;

    for (let i = 0; i < series.length; i += 1) {
      const s = series[i];
      const percentual = total > 0 ? (s.valor / total) * 100 : s.percentual;
      const angle = (percentual / 100) * Math.PI * 2;
      const color = this.paletteWorkbook[i % this.paletteWorkbook.length];
      const middle = current + angle / 2;

      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, radius, current, current + angle);
      ctx.closePath();
      ctx.fillStyle = color;
      ctx.fill();

      if (percentual > 0) {
        const linha1 = this.formatarNumero(s.valor, 1);
        const linha2 = `${this.formatarNumero(s.percentual, 1)}%`;
        const textW = Math.max(ctx.measureText(linha1).width, ctx.measureText(linha2).width);
        const boxW = textW + 12;
        const boxH = Math.max(24, Math.round(labelFont * 2.25));
        const outsidePreferred = percentual < 12;

        const radiusOptions = [radius * 0.62, radius * 0.52, radius * 0.72, radius * 0.42];
        const offsetOptions = [0, -angle * 0.22, angle * 0.22, -angle * 0.35, angle * 0.35];

        let placed = false;
        if (!outsidePreferred) {
          for (const r of radiusOptions) {
            for (const delta of offsetOptions) {
              const ang = middle + delta;
              const lx = cx + Math.cos(ang) * r;
              const ly = cy + Math.sin(ang) * r;
              const dist = Math.hypot(lx - cx, ly - cy);
              const maxDist = radius - Math.max(boxW, boxH) * 0.25;
              if (dist > maxDist) continue;
              if (collide(lx, ly, boxW, boxH)) continue;
              labels.push({ linhas: [linha1, linha2], x: lx, y: ly, w: boxW, h: boxH });
              placed = true;
              break;
            }
            if (placed) break;
          }
        }

        if (!placed) {
          const side: "left" | "right" = Math.cos(middle) >= 0 ? "right" : "left";
          labelsExternos.push({
            linhas: [linha1, linha2],
            x: side === "right" ? plotRight + 8 : plotLeft - 8,
            y: cy + Math.sin(middle) * (radius * 0.95),
            w: boxW,
            h: boxH,
            side,
            anchorX: cx + Math.cos(middle) * (radius * 0.98),
            anchorY: cy + Math.sin(middle) * (radius * 0.98),
          });
        }
      }

      current += angle;
    }

    const ajustarExternos = (side: "left" | "right") => {
      const itens = labelsExternos.filter((l) => l.side === side).sort((a, b) => a.y - b.y);
      const minY = plotTop + 6;
      const maxY = plotBottom - 6;
      for (let i = 0; i < itens.length; i += 1) {
        if (i > 0) {
          const minGap = Math.max(itens[i].h, itens[i - 1].h) + 4;
          if (itens[i].y - itens[i - 1].y < minGap) itens[i].y = itens[i - 1].y + minGap;
        }
      }
      for (let i = itens.length - 1; i >= 0; i -= 1) {
        const top = minY + itens[i].h / 2;
        const bottom = maxY - itens[i].h / 2;
        if (itens[i].y > bottom) itens[i].y = bottom;
        if (itens[i].y < top) itens[i].y = top;
        if (i > 0) {
          const minGap = Math.max(itens[i].h, itens[i - 1].h) + 4;
          if (itens[i].y - itens[i - 1].y < minGap) itens[i - 1].y = itens[i].y - minGap;
        }
      }
      for (const item of itens) {
        const top = minY + item.h / 2;
        const bottom = maxY - item.h / 2;
        if (item.y < top) item.y = top;
        if (item.y > bottom) item.y = bottom;
        if (side === "right") {
          const minX = cx + radius + 6;
          const maxX = legendX - item.w - 10;
          if (maxX > minX) {
            item.x = Math.max(minX, Math.min(maxX, item.anchorX + 10));
          } else {
            item.x = Math.max(6, Math.min(width - item.w - 6, item.anchorX + 6));
          }
        } else {
          const minRight = Math.max(item.w + 6, plotLeft + item.w + 2);
          const maxRight = cx - radius - 6;
          if (maxRight > minRight) {
            item.x = Math.max(minRight, Math.min(maxRight, item.anchorX - 10));
          } else {
            item.x = Math.max(item.w + 6, Math.min(width - 6, item.anchorX - 6));
          }
        }
      }
    };

    ajustarExternos("left");
    ajustarExternos("right");

    const drawRoundRect = (x: number, y: number, w: number, h: number, r: number) => {
      ctx.beginPath();
      ctx.moveTo(x + r, y);
      ctx.lineTo(x + w - r, y);
      ctx.quadraticCurveTo(x + w, y, x + w, y + r);
      ctx.lineTo(x + w, y + h - r);
      ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
      ctx.lineTo(x + r, y + h);
      ctx.quadraticCurveTo(x, y + h, x, y + h - r);
      ctx.lineTo(x, y + r);
      ctx.quadraticCurveTo(x, y, x + r, y);
      ctx.closePath();
    };

    ctx.font = `bold ${labelFont}px DM Sans, Segoe UI, sans-serif`;
    for (const label of labels) {
      const boxW = label.w;
      const boxH = label.h;
      let boxX = label.x - boxW / 2;
      const boxY = label.y - boxH / 2;
      if (boxX < 4) boxX = 4;
      if (boxX + boxW > width - 4) boxX = width - boxW - 4;

      drawRoundRect(boxX, boxY, boxW, boxH, 6);
      ctx.fillStyle = "rgba(255,255,255,0.95)";
      ctx.fill();
      ctx.strokeStyle = "#d1d5db";
      ctx.lineWidth = 1;
      ctx.stroke();

      ctx.fillStyle = "#102a43";
      ctx.fillText(label.linhas[0], boxX + 5, boxY + Math.round(boxH * 0.45));
      ctx.fillText(label.linhas[1], boxX + 5, boxY + Math.round(boxH * 0.82));
    }

    ctx.strokeStyle = "#64748b";
    ctx.lineWidth = 1;
    for (const label of labelsExternos) {
      const boxW = label.w;
      const boxH = label.h;
      const boxX = label.side === "right" ? label.x : label.x - boxW;
      const boxY = label.y - boxH / 2;

      const toX = label.side === "right" ? boxX : boxX + boxW;
      const toY = label.y;
      ctx.beginPath();
      ctx.moveTo(label.anchorX, label.anchorY);
      ctx.lineTo(toX, toY);
      ctx.stroke();

      drawRoundRect(boxX, boxY, boxW, boxH, 6);
      ctx.fillStyle = "rgba(255,255,255,0.95)";
      ctx.fill();
      ctx.strokeStyle = "#d1d5db";
      ctx.stroke();
      ctx.fillStyle = "#102a43";
      ctx.fillText(label.linhas[0], boxX + 5, boxY + Math.round(boxH * 0.45));
      ctx.fillText(label.linhas[1], boxX + 5, boxY + Math.round(boxH * 0.82));
      ctx.strokeStyle = "#64748b";
    }

    let y = Math.max(58, Math.round(height * 0.22));
    const swatch = Math.max(12, Math.round(legendFont * 0.95));
    const legendTextX = legendX + swatch + 8;
    const legendTextMaxW = Math.max(48, width - legendTextX - 6);

    const wrapLegendText = (text: string): string[] => {
      const words = text.split(/\s+/).filter(Boolean);
      const lines: string[] = [];
      let current = "";
      for (const word of words) {
        const test = current ? `${current} ${word}` : word;
        if (ctx.measureText(test).width <= legendTextMaxW || !current) {
          current = test;
        } else {
          lines.push(current);
          current = word;
        }
      }
      if (current) lines.push(current);
      return lines.slice(0, 2);
    };

    for (let i = 0; i < series.length; i += 1) {
      const s = series[i];
      const color = this.paletteWorkbook[i % this.paletteWorkbook.length];

      ctx.fillStyle = color;
      ctx.fillRect(legendX, y - swatch + 2, swatch, swatch);
      ctx.fillStyle = "#102a43";
      ctx.font = `${legendFont}px DM Sans, Segoe UI, sans-serif`;
      const legendLines = wrapLegendText(`${s.label}`);
      ctx.fillText(legendLines[0] ?? "", legendTextX, y + 1);
      if (legendLines[1]) {
        y += Math.max(14, Math.round(legendFont * 1.25));
        ctx.fillText(legendLines[1], legendTextX, y + 1);
      }
      y += Math.max(20, Math.round(legendFont * 1.45));
      if (y > height - 18) break;
    }

    return this.dataUrlParaBytes(canvas.toDataURL("image/png"));
  }

  private agruparGraficosPorTrecho(graficos: WorkbookGrafico[]): WorkbookTrechoGroup[] {
    const porTrecho = new Map<string, WorkbookGrafico[]>();
    for (const g of graficos) {
      const trecho = (g.trecho ?? "Trecho nao informado").trim() || "Trecho nao informado";
      const lista = porTrecho.get(trecho) ?? [];
      lista.push(g);
      porTrecho.set(trecho, lista);
    }

    return Array.from(porTrecho.entries())
      .map(([trecho, lista]) => ({
        trecho,
        graficos: [...lista].sort((a, b) => a.titulo.localeCompare(b.titulo, "pt-BR")),
      }))
      .sort((a, b) => a.trecho.localeCompare(b.trecho, "pt-BR"));
  }

  private chaveTrechoTemplate(value: string): string {
    return this.normalizarTextoRelatorio(value)
      .replace(/^[\d.]+\s+/g, "")
      .replace(/\s*\(.*?\)\s*/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  private extrairTrechoTemplate(texto: string): string | null {
    const semPrefixo = texto.replace(/^[\d.]+\s+/, "").trim();
    const idx = semPrefixo.toUpperCase().indexOf("RODOVIA ");
    if (idx === -1) return null;
    const trecho = semPrefixo.slice(idx).trim();
    return trecho || null;
  }

  private ordemTrechosTemplate(tipo: "PAV" | "NAO_PAV"): string[] {
    if (this.templateEstruturaHeadings.length === 0) return [];

    const resultado: string[] = [];
    let contexto: "PAV" | "NAO_PAV" | null = null;

    for (const heading of this.templateEstruturaHeadings) {
      const normalizado = this.normalizarTextoRelatorio(heading.texto);
      if (normalizado.includes("rodovias pavimentadas") && !normalizado.includes("nao pavimentadas")) {
        contexto = "PAV";
        continue;
      }
      if (normalizado.includes("rodovias nao pavimentadas")) {
        contexto = "NAO_PAV";
        continue;
      }
      if (contexto !== tipo) continue;

      const trecho = this.extrairTrechoTemplate(heading.texto);
      if (!trecho) continue;
      resultado.push(trecho);
    }

    return resultado;
  }

  private ordenarGruposPorTemplate(tipo: "PAV" | "NAO_PAV", grupos: WorkbookTrechoGroup[]): WorkbookTrechoGroup[] {
    const ordemTemplate = this.ordemTrechosTemplate(tipo);
    if (ordemTemplate.length === 0) return grupos;

    const pendentes = [...grupos];
    const ordenados: WorkbookTrechoGroup[] = [];

    const matchKey = (a: string, b: string): boolean => {
      const ka = this.chaveTrechoTemplate(a);
      const kb = this.chaveTrechoTemplate(b);
      return ka === kb || ka.includes(kb) || kb.includes(ka);
    };

    for (const trechoTemplate of ordemTemplate) {
      const idx = pendentes.findIndex((g) => matchKey(g.trecho, trechoTemplate));
      if (idx === -1) continue;
      const [grupo] = pendentes.splice(idx, 1);
      ordenados.push({ ...grupo, tituloTemplate: trechoTemplate });
    }

    return [...ordenados, ...pendentes];
  }

  private async renderizarBlocoGraficosTrecho(
    tipo: "PAV" | "NAO_PAV",
    trecho: string,
    graficos: WorkbookGrafico[],
  ): Promise<{ data: Uint8Array; width: number; height: number }> {
    const tileW = 420;
    const tileH = 290;
    const cols = 2;
    const gap = 20;
    const pad = 18;
    const headerH = 64;
    const rows = Math.max(1, Math.ceil(graficos.length / cols));
    const width = pad * 2 + cols * tileW + gap * (cols - 1);
    const height = headerH + pad + rows * tileH + (rows - 1) * gap + pad;

    const canvas = document.createElement("canvas");
    const dpr = 2;
    canvas.width = width * dpr;
    canvas.height = height * dpr;
    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Falha ao inicializar bloco grafico do relatorio.");
    ctx.scale(dpr, dpr);

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);

    ctx.fillStyle = "#1f2933";
    ctx.font = "bold 22px DM Sans, Segoe UI, sans-serif";
    const tituloTrecho = this.isTrechoNaoInformado(trecho) ? this.resumoAnalisePorTipo(tipo) : trecho;
    ctx.fillText(`${tipo} - ${tituloTrecho}`.slice(0, 82), pad, 38);

    for (let i = 0; i < graficos.length; i += 1) {
      const g = graficos[i];
      const row = Math.floor(i / cols);
      const col = i % cols;
      const x = pad + col * (tileW + gap);
      const y = headerH + pad + row * (tileH + gap);

      const bytes = await this.renderizarPizzaWorkbook(g.titulo, g.series, { width: tileW, height: tileH });
      const img = await this.bytesParaImagem(bytes);
      ctx.drawImage(img, x, y, tileW, tileH);
    }

    return { data: this.dataUrlParaBytes(canvas.toDataURL("image/png")), width, height };
  }

  private selecionarPrimeiroGraficoPorTitulo(graficos: WorkbookGrafico[]): WorkbookGrafico[] {
    const ordenarAba = (a: WorkbookGrafico, b: WorkbookGrafico): number => {
      if (a.aba === b.aba) return a.ordem - b.ordem;
      if (a.aba === "TT" && b.aba !== "TT") return 1;
      if (b.aba === "TT" && a.aba !== "TT") return -1;

      const na = Number(a.aba);
      const nb = Number(b.aba);
      if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
      if (Number.isFinite(na)) return -1;
      if (Number.isFinite(nb)) return 1;
      return a.aba.localeCompare(b.aba, "pt-BR");
    };

    const porTitulo = new Map<string, WorkbookGrafico[]>();
    for (const g of graficos) {
      const titulo = g.titulo.trim().toUpperCase();
      const lista = porTitulo.get(titulo) ?? [];
      lista.push(g);
      porTitulo.set(titulo, lista);
    }

    const selecionados: WorkbookGrafico[] = [];
    for (const lista of porTitulo.values()) {
      const ordenada = [...lista].sort(ordenarAba);
      const preferido = ordenada.find((g) => g.aba === "TT") ?? ordenada[0];
      if (preferido) selecionados.push(preferido);
    }

    return selecionados.sort((a, b) => a.titulo.localeCompare(b.titulo, "pt-BR"));
  }

  private canonicalTexto(value: string): string {
    return this.normalizarTextoRelatorio(value).replace(/[^a-z0-9]/g, "");
  }

  private decodeXmlText(value: string): string {
    return value
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#10;/g, "\n")
      .replace(/&#13;/g, "\r");
  }

  private extractParagraphTextXml(paraXml: string): string {
    const chunks = [...paraXml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)].map((m) => this.decodeXmlText(m[1]));
    return chunks.join(" ").replace(/\s+/g, " ").trim();
  }

  private detectTipoContextoTemplate(text: string): "PAV" | "NAO_PAV" | undefined {
    const n = this.normalizarTextoRelatorio(text);
    if (n.includes("rodovias nao pavimentadas")) return "NAO_PAV";
    if (n.includes("rodovias pavimentadas") && !n.includes("nao pavimentadas")) return "PAV";
    if (n.includes("1 6 1") && n.includes("rodovias") && n.includes("pavimentadas")) return "PAV";
    if (n.includes("1 6 2") && n.includes("rodovias") && n.includes("nao pavimentadas")) return "NAO_PAV";
    return undefined;
  }

  private isTrechoHeadingTemplate(text: string): boolean {
    return /RODOVIA\s+[A-Z]{2}-?\d+/i.test(text);
  }

  private sanitizeTrechoHeadingTemplate(text: string): string {
    const rodoviaStart = text.search(/RODOVIA\s+[A-Z]{2}-?\d+/i);
    const base = rodoviaStart >= 0 ? text.slice(rodoviaStart) : text;
    return base
      .replace(/PAGEREF\s+_Toc\d+\s+\\h\s*\d*$/i, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  private isGraficoCaptionTemplate(text: string): boolean {
    const n = this.normalizarTextoRelatorio(text);
    return n.includes("avaliacao do consorcio supervisor") && n.includes("condicoes de pista") && n.includes("extrapista");
  }

  private updateRelationshipTargetXml(relsXml: string, rid: string, newTarget: string): string {
    const rx = new RegExp(`(<Relationship\\b[^>]*\\bId="${rid}"[^>]*\\bTarget=")(?:[^"]+)("[^>]*>)`);
    return relsXml.replace(rx, `$1${newTarget}$2`);
  }

  private pxToEmu(px: number): number {
    return Math.max(1, Math.round(px * 9525));
  }

  private cmToEmu(cm: number): number {
    return Math.max(1, Math.round(cm * 360000));
  }

  private atualizarExtensaoImagemNoDocxTamanhoFixo(
    documentXml: string,
    rid: string,
    larguraCm: number,
    alturaCm: number,
  ): string {
    const atualizarSegmento = (segment: string): string => {
      const targetHeightEmu = this.cmToEmu(alturaCm);
      const targetWidthEmu = this.cmToEmu(larguraCm);

      let out = segment;
      out = out.replace(
        /(<wp:extent\b[^>]*\bcx=")(\d+)("[^>]*\bcy=")(\d+)("[^>]*\/>)/,
        `$1${targetWidthEmu}$3${targetHeightEmu}$5`,
      );
      out = out.replace(
        /(<pic:spPr\b[\s\S]*?<a:xfrm\b[^>]*>[\s\S]*?<a:off\b[^>]*\bx=")(\-?\d+)("[^>]*\by=")(\-?\d+)("[^>]*\/>)/,
        (_, p1: string, _x: string, p3: string, _y: string, p5: string) => `${p1}0${p3}0${p5}`,
      );
      out = out.replace(
        /(<pic:spPr\b[\s\S]*?<a:xfrm\b[^>]*>[\s\S]*?<a:ext\b[^>]*\bcx=")(\d+)("[^>]*\bcy=")(\d+)("[^>]*\/>)/,
        `$1${targetWidthEmu}$3${targetHeightEmu}$5`,
      );
      out = out.replace(/<a:srcRect\b[^>]*\/>/g, '<a:srcRect l="0" t="0" r="0" b="0"/>');
      return out;
    };

    const updateByTag = (xml: string, tag: "inline" | "anchor"): string => {
      const pattern = new RegExp(`<wp:${tag}\\b[\\s\\S]*?<a:blip[^>]*r:embed="${rid}"[^>]*>[\\s\\S]*?<\\/wp:${tag}>`, "g");
      return xml.replace(pattern, (segment) => atualizarSegmento(segment));
    };

    return updateByTag(updateByTag(documentXml, "inline"), "anchor");
  }

  private pickGrupoTrechoTemplate(
    gruposByKey: Map<string, WorkbookTrechoGroup>,
    tipo: "PAV" | "NAO_PAV",
    trecho: string,
  ): { grupo: WorkbookTrechoGroup | undefined; key: string | undefined } {
    const key = `${tipo}|${this.canonicalTexto(trecho)}`;
    const exact = gruposByKey.get(key);
    if (exact) return { grupo: exact, key };

    return { grupo: undefined, key: undefined };
  }

  private pickGrupoTrechoTemplateComFallback(
    gruposByKey: Map<string, WorkbookTrechoGroup>,
    tipoPreferido: "PAV" | "NAO_PAV",
    trecho: string,
  ): { grupo: WorkbookTrechoGroup | undefined; key: string | undefined; tipoUsado: "PAV" | "NAO_PAV" | undefined } {
    const primary = this.pickGrupoTrechoTemplate(gruposByKey, tipoPreferido, trecho);
    if (primary.grupo) return { ...primary, tipoUsado: tipoPreferido };
    return { grupo: undefined, key: undefined, tipoUsado: undefined };
  }


  private substituirCamposCapaTemplate(documentXml: string): string {
    const regiaoTexto = `REGIAO ${String(this.regiao).padStart(2, "0")} DE CONSERVACAO`;
    const mesAnoTexto = `${this.mesLabel(this.mes).toUpperCase()} DE ${this.ano}`;

    let updated = documentXml.replace(/REGI[ÃA]O\s*\d{1,2}\s*DE\s*CONSERVA[CÇ][ÃA]O/gi, regiaoTexto);
    updated = updated.replace(
      /(JANEIRO|FEVEREIRO|MAR[ÇC]O|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)\s+DE\s+\d{4}/gi,
      mesAnoTexto,
    );

    return updated;
  }

  private async gerarDocxRegionalViaTemplate(
    template: File,
    workbookPav: WorkbookGrafico[],
    workbookNaoPav: WorkbookGrafico[],
  ): Promise<Blob> {
    const zip = await JSZip.loadAsync(await template.arrayBuffer());
    let documentXml = await zip.file("word/document.xml")?.async("string");
    let relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");
    if (!documentXml || !relsXml) {
      throw new Error("Template DOCX invalido: document.xml ou relationships ausentes.");
    }

    documentXml = this.substituirCamposCapaTemplate(documentXml);

    const gruposByKey = new Map<string, WorkbookTrechoGroup>();
    const gruposPav = this.agruparGraficosPorTrecho(workbookPav);
    const gruposNaoPav = this.agruparGraficosPorTrecho(workbookNaoPav);
    for (const grupo of gruposPav) {
      gruposByKey.set(`PAV|${this.canonicalTexto(grupo.trecho)}`, grupo);
    }
    for (const grupo of gruposNaoPav) {
      gruposByKey.set(`NAO_PAV|${this.canonicalTexto(grupo.trecho)}`, grupo);
    }

    const blocoCache = new Map<string, { data: Uint8Array; width: number; height: number }>();

    const paragraphMatches = [...documentXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)];
    let tipoAtual: "PAV" | "NAO_PAV" = "PAV";
    let trechoAtual = "";
    let aguardandoGrafico = false;
    let emResumoGeral = false;
    let substituidos = 0;

    for (const paraMatch of paragraphMatches) {
      const paraXml = paraMatch[0];
      const text = this.extractParagraphTextXml(paraXml);
      const normalizado = text ? this.normalizarTextoRelatorio(text) : "";

      if (normalizado.includes("resumo de analise do geral")) {
        emResumoGeral = true;
      }

      const tipoDetectado = text ? this.detectTipoContextoTemplate(text) : undefined;
      if (tipoDetectado) tipoAtual = tipoDetectado;

      if (text && this.isTrechoHeadingTemplate(text)) {
        trechoAtual = this.sanitizeTrechoHeadingTemplate(text);
        emResumoGeral = false;
      }
      if (text && this.isGraficoCaptionTemplate(text)) aguardandoGrafico = true;

      const embeds = [...paraXml.matchAll(/<a:blip[^>]*r:embed="([^"]+)"/g)].map((m) => m[1]);
      if (!embeds.length || !aguardandoGrafico) continue;

      const trechoBusca = emResumoGeral ? "Trecho nao informado" : trechoAtual;
      const { grupo, key, tipoUsado } = this.pickGrupoTrechoTemplateComFallback(gruposByKey, tipoAtual, trechoBusca);
      if (!grupo || !key || !tipoUsado) {
        aguardandoGrafico = false;
        continue;
      }

      tipoAtual = tipoUsado;

      let bloco = blocoCache.get(key);
      if (!bloco) {
        const bloco = await this.renderizarBlocoGraficosTrecho(tipoUsado, grupo.tituloTemplate ?? grupo.trecho, grupo.graficos);
        blocoCache.set(key, bloco);
      }
      bloco = blocoCache.get(key);
      if (!bloco) {
        aguardandoGrafico = false;
        continue;
      }

      for (const rid of embeds) {
        const mediaName = `grafico_atualizado_${String(substituidos + 1).padStart(3, "0")}.png`;
        zip.file(`word/media/${mediaName}`, bloco.data);
        relsXml = this.updateRelationshipTargetXml(relsXml, rid, `media/${mediaName}`);
        const alturaCm = tipoUsado === "PAV" ? this.relatorioAlturaPavCm : this.relatorioAlturaNaoPavCm;
        const larguraCm = this.relatorioLarguraBlocoCm;
        documentXml = this.atualizarExtensaoImagemNoDocxTamanhoFixo(
          documentXml,
          rid,
          larguraCm,
          alturaCm,
        );
        substituidos += 1;
      }
      aguardandoGrafico = false;
    }

    zip.file("word/document.xml", documentXml);
    zip.file("word/_rels/document.xml.rels", relsXml);
    return zip.generateAsync({ type: "blob" });
  }

  private tabela2Colunas(linhas: Array<[string, string]>): Table {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: linhas.map(
        ([k, v]) =>
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: k, bold: true })] })] }),
              new TableCell({ children: [new Paragraph(v)] }),
            ],
          }),
      ),
    });
  }

  async baixarDocxCompetencia(): Promise<void> {
    if (!this.sessaoAtual?.token) return;
    this.salvarDimensoesRelatorio();

    try {
      const [payload, graficos, inconsistencias, base, workbookPavRaw, workbookNaoPavRaw] = await Promise.all([
        this.client.query(api.trechos.gerarPayloadRelatorio, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        this.client.query(api.trechos.obterGraficosCompetencia, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        this.client.query(api.trechos.obterInconsistenciasImportacao, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        this.client.query(api.trechos.obterBaseConsolidada, {
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
        }),
        this.client.query(api.workbook.listarGraficosWorkbook, {
          sessionToken: this.tokenSessao(),
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
          tipoFonte: "PAV",
        }),
        this.client.query(api.workbook.listarGraficosWorkbook, {
          sessionToken: this.tokenSessao(),
          regiao: this.regiao,
          ano: this.ano,
          mes: this.mes,
          tipoFonte: "NAO_PAV",
        }),
      ]);

      const trechos = (base as any).trechos as Array<any>;
      const importacoes = (base as any).importacoes as Array<any>;

      const pav = trechos.filter((t) => t.tipoFonte === "PAV");
      const naoPav = trechos.filter((t) => t.tipoFonte === "NAO_PAV");

      const pavKm = pav.reduce((acc, t) => acc + (t.extKm ?? 0), 0);
      const naoPavKm = naoPav.reduce((acc, t) => acc + (t.extKm ?? 0), 0);

      const workbookPav = workbookPavRaw as WorkbookGrafico[];
      const workbookNaoPav = workbookNaoPavRaw as WorkbookGrafico[];

      if (this.templateDocxBase) {
        const blobTemplate = await this.gerarDocxRegionalViaTemplate(this.templateDocxBase, workbookPav, workbookNaoPav);
        const urlTemplate = URL.createObjectURL(blobTemplate);
        const aTemplate = document.createElement("a");
        aTemplate.href = urlTemplate;
        aTemplate.download = `relatorio_regiao_${this.regiao}_${this.competencia()}.docx`;
        aTemplate.click();
        URL.revokeObjectURL(urlTemplate);
        return;
      }

      const titulosUnicosPav = new Set((workbookPavRaw as WorkbookGrafico[]).map((g) => g.titulo.trim().toUpperCase())).size;
      const titulosUnicosNaoPav = new Set((workbookNaoPavRaw as WorkbookGrafico[]).map((g) => g.titulo.trim().toUpperCase())).size;
      const tituloApresentacao = this.headingTemplateOuPadrao("1. Apresentacao", ["apresentacao"]);
      const tituloVistorias = this.headingTemplateOuPadrao("3. Vistorias Rotineiras no Periodo", ["vistorias", "periodo"]);
      const tituloPavimentadas = this.headingTemplateOuPadrao("3.1 Rodovias Pavimentadas", ["rodovias pavimentadas"]);
      const tituloNaoPavimentadas = this.headingTemplateOuPadrao("3.2 Rodovias Nao Pavimentadas", ["rodovias nao pavimentadas"]);
      const tituloEncerramento = this.headingTemplateOuPadrao("6. Conclusoes Automaticas", ["encerramento"]);

      const children: Array<Paragraph | Table> = [
        new Paragraph({
          text: "ESTADO DO TOCANTINS",
          alignment: AlignmentType.CENTER,
          heading: HeadingLevel.HEADING_1,
        }),
        new Paragraph({
          text: "AGENCIA DE TRANSPORTES, OBRAS E INFRAESTRUTURA - AGETO",
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({ text: "" }),
        new Paragraph({
          text: "PRODUTO 02 - RELATORIO DE ACOMPANHAMENTO TECNICO / AMBIENTAL",
          alignment: AlignmentType.CENTER,
          heading: HeadingLevel.HEADING_2,
        }),
        new Paragraph({
          text: `REGIAO ${String(this.regiao).padStart(2, "0")} DE CONSERVACAO`,
          alignment: AlignmentType.CENTER,
          heading: HeadingLevel.HEADING_2,
        }),
        new Paragraph({
          text: `${this.mesLabel(this.mes).toUpperCase()} DE ${this.ano}`,
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({ children: [new PageBreak()] }),

        new Paragraph({ text: "SUMARIO", heading: HeadingLevel.HEADING_1 }),
        new TableOfContents("", { headingStyleRange: "1-3", hyperlink: true }),
        new Paragraph({ children: [new PageBreak()] }),

        new Paragraph({ text: tituloApresentacao, heading: HeadingLevel.HEADING_1 }),
        new Paragraph(
          `Este relatorio sintetiza os resultados da competencia ${this.competencia()} para a Regiao ${this.regiao}, com base nas importacoes de planilhas, auditoria operacional e graficos complementares do workbook.`,
        ),
        new Paragraph({ text: "" }),

        new Paragraph({ text: "2. Indicadores Gerais", heading: HeadingLevel.HEADING_1 }),
        this.tabela2Colunas([
          ["Competencia", this.competencia()],
          ["Total de trechos", String((graficos as any).kpis.totalTrechos)],
          ["Total de extensao (km)", this.formatarNumero((graficos as any).kpis.totalKm, 2)],
          ["Programados no mes", String((graficos as any).kpis.programadosNoMes)],
          ["Nao programados no mes", String((graficos as any).kpis.naoProgramadosNoMes)],
          ["Percentual programados", `${(payload as any).graficos.kpis.percentualProgramados}%`],
        ]),
        new Paragraph({ text: "" }),

        new Paragraph({ text: tituloVistorias, heading: HeadingLevel.HEADING_1 }),
        new Paragraph({ text: tituloPavimentadas, heading: HeadingLevel.HEADING_2 }),
        this.tabela2Colunas([
          ["Trechos (linhas)", String(pav.length)],
          ["Trechos unicos", String(new Set(pav.map((t) => t.trecho)).size)],
          ["Extensao total (km)", this.formatarNumero(pavKm, 2)],
          ["Graficos workbook (total)", String((workbookPavRaw as WorkbookGrafico[]).length)],
          ["Abas com grafico", String(new Set((workbookPavRaw as WorkbookGrafico[]).map((g) => g.aba)).size)],
        ]),
        new Paragraph({ text: "" }),

        new Paragraph({ text: tituloNaoPavimentadas, heading: HeadingLevel.HEADING_2 }),
        this.tabela2Colunas([
          ["Trechos (linhas)", String(naoPav.length)],
          ["Trechos unicos", String(new Set(naoPav.map((t) => t.trecho)).size)],
          ["Extensao total (km)", this.formatarNumero(naoPavKm, 2)],
          ["Graficos workbook (total)", String((workbookNaoPavRaw as WorkbookGrafico[]).length)],
          ["Abas com grafico", String(new Set((workbookNaoPavRaw as WorkbookGrafico[]).map((g) => g.aba)).size)],
        ]),
        new Paragraph({ text: "" }),

        new Paragraph({ text: "4. Controle de Importacoes e Inconsistencias", heading: HeadingLevel.HEADING_1 }),
        this.tabela2Colunas([
          ["Total de importacoes", String((inconsistencias as any).resumo.totalImportacoes)],
          ["Importacoes com erro", String((inconsistencias as any).resumo.importacoesComErro)],
          ["Total de erros", String((inconsistencias as any).resumo.totalErros)],
          ["Ultima importacao", this.formatarDataHora(importacoes[0]?.finalizadoEm ?? importacoes[0]?.iniciadoEm)],
        ]),
      ];

      children.push(new Paragraph({ text: "" }));
      children.push(new Paragraph({ text: "5. Graficos do Workbook", heading: HeadingLevel.HEADING_1 }));
      children.push(
        new Paragraph(
          "Para padronizacao da emissao automatica, esta secao inclui todos os graficos extraidos do workbook, mantendo os valores e percentuais por aba.",
        ),
      );
      children.push(
        this.tabela2Colunas([
          ["PAV - graficos brutos", String((workbookPavRaw as WorkbookGrafico[]).length)],
          ["PAV - titulos unicos", String(titulosUnicosPav)],
          ["PAV - graficos inseridos", String(workbookPav.length)],
          ["NAO_PAV - graficos brutos", String((workbookNaoPavRaw as WorkbookGrafico[]).length)],
          ["NAO_PAV - titulos unicos", String(titulosUnicosNaoPav)],
          ["NAO_PAV - graficos inseridos", String(workbookNaoPav.length)],
        ]),
      );

      children.push(new Paragraph({ text: "" }));
      children.push(new Paragraph({ text: "5.1 Rodovias Pavimentadas", heading: HeadingLevel.HEADING_2 }));
      if (workbookPav.length === 0) {
        children.push(
          new Paragraph(
            "Nenhum grafico de workbook foi encontrado para Rodovias Pavimentadas nesta competencia. Verifique se a importacao complementar PAV foi executada.",
          ),
        );
      } else {
        const pavPorTrecho = this.ordenarGruposPorTemplate("PAV", this.agruparGraficosPorTrecho(workbookPav));
        for (const grupo of pavPorTrecho) {
          const bloco = await this.renderizarBlocoGraficosTrecho("PAV", grupo.trecho, grupo.graficos);
          const largura = Math.min(620, bloco.width);
          const altura = Math.round((bloco.height / bloco.width) * largura);
          children.push(new Paragraph({ text: grupo.tituloTemplate ?? `Trecho ${grupo.trecho}`, heading: HeadingLevel.HEADING_3 }));
          children.push(
            new Paragraph({
              children: [new ImageRun({ type: "png", data: bloco.data, transformation: { width: largura, height: altura } })],
            }),
          );
        }
      }

      children.push(new Paragraph({ text: "" }));
      children.push(new Paragraph({ text: "5.2 Rodovias Nao Pavimentadas", heading: HeadingLevel.HEADING_2 }));
      if (workbookNaoPav.length === 0) {
        children.push(
          new Paragraph(
            "Nenhum grafico de workbook foi encontrado para Rodovias Nao Pavimentadas nesta competencia. Verifique se a importacao complementar NAO_PAV foi executada.",
          ),
        );
      } else {
        const naoPavPorTrecho = this.ordenarGruposPorTemplate("NAO_PAV", this.agruparGraficosPorTrecho(workbookNaoPav));
        for (const grupo of naoPavPorTrecho) {
          const bloco = await this.renderizarBlocoGraficosTrecho("NAO_PAV", grupo.trecho, grupo.graficos);
          const largura = Math.min(620, bloco.width);
          const altura = Math.round((bloco.height / bloco.width) * largura);
          children.push(new Paragraph({ text: grupo.tituloTemplate ?? `Trecho ${grupo.trecho}`, heading: HeadingLevel.HEADING_3 }));
          children.push(
            new Paragraph({
              children: [new ImageRun({ type: "png", data: bloco.data, transformation: { width: largura, height: altura } })],
            }),
          );
        }
      }

      children.push(new Paragraph({ children: [new PageBreak()] }));
      children.push(new Paragraph({ text: tituloEncerramento, heading: HeadingLevel.HEADING_1 }));
      for (const obs of (payload as any).observacoesAutomaticas as string[]) {
        children.push(new Paragraph({ text: obs, bullet: { level: 0 } }));
      }

      const doc = new Document({ sections: [{ children }] });
      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `relatorio_regiao_${this.regiao}_${this.competencia()}.docx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      this.erro = e instanceof Error ? e.message : String(e);
    }
  }

  async baixarMarkdownConsolidado(): Promise<void> {
    this.gerandoConsolidado = true;
    this.erro = "";

    try {
      const resultados = await Promise.all(
        this.regioes.map(async (regiao) => {
          const [graficos, inconsistencias] = await Promise.all([
            this.client.query(api.trechos.obterGraficosCompetencia, {
              regiao,
              ano: this.ano,
              mes: this.mes,
            }),
            this.client.query(api.trechos.obterInconsistenciasImportacao, {
              regiao,
              ano: this.ano,
              mes: this.mes,
            }),
          ]);

          return {
            regiao,
            totalTrechos: graficos.kpis.totalTrechos,
            totalKm: graficos.kpis.totalKm,
            totalImportacoes: inconsistencias.resumo.totalImportacoes,
            importacoesComErro: inconsistencias.resumo.importacoesComErro,
            totalErros: inconsistencias.resumo.totalErros,
            porTipoFonte: graficos.series.porTipoFonte as FonteResumo[],
          };
        }),
      );

      const totais = resultados.reduce(
        (acc, item) => {
          acc.totalTrechos += item.totalTrechos;
          acc.totalKm += item.totalKm;
          acc.totalImportacoes += item.totalImportacoes;
          acc.importacoesComErro += item.importacoesComErro;
          acc.totalErros += item.totalErros;
          return acc;
        },
        {
          totalTrechos: 0,
          totalKm: 0,
          totalImportacoes: 0,
          importacoesComErro: 0,
          totalErros: 0,
        },
      );

      const linhas = [
        "# Relatorio Tecnico Consolidado",
        "",
        `- Competencia: ${this.competencia()}`,
        `- Regioes: ${this.regioes.join(", ")}`,
        `- Gerado em: ${new Date().toISOString()}`,
        "",
        "## Totais Consolidados",
        `- Total de trechos: ${totais.totalTrechos}`,
        `- Total de extensao (km): ${totais.totalKm.toFixed(2)}`,
        `- Total de importacoes: ${totais.totalImportacoes}`,
        `- Importacoes com erro: ${totais.importacoesComErro}`,
        `- Total de erros: ${totais.totalErros}`,
        "",
        "## Resumo por Regiao",
        ...resultados.flatMap((item) => [
          `### Regiao ${item.regiao}`,
          `- Total de trechos: ${item.totalTrechos}`,
          `- Total de extensao (km): ${item.totalKm.toFixed(2)}`,
          `- Total de importacoes: ${item.totalImportacoes}`,
          `- Importacoes com erro: ${item.importacoesComErro}`,
          `- Total de erros: ${item.totalErros}`,
          "",
          "- Distribuicao por tipo de fonte:",
          ...item.porTipoFonte.map(
            (fonte) => `  - ${fonte.tipoFonte}: ${fonte.totalTrechos} trechos / ${fonte.totalKm.toFixed(2)} km`,
          ),
          "",
          `- Relatorio detalhado: relatorio_regiao_${item.regiao}_${this.competencia()}.md`,
          "",
        ]),
      ];

      const blob = new Blob([linhas.join("\n")], { type: "text/markdown;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `relatorio_consolidado_${this.competencia()}.md`;
      a.click();
      URL.revokeObjectURL(url);

      this._ultimasLinhasConsolidado = linhas;
    } catch (e) {
      this.erro = e instanceof Error ? e.message : String(e);
    } finally {
      this.gerandoConsolidado = false;
    }
  }

  private _ultimasLinhasConsolidado: string[] | null = null;

  async baixarDocxConsolidado(): Promise<void> {
    if (!this._ultimasLinhasConsolidado) {
      await this.baixarMarkdownConsolidado();
    }
    if (!this._ultimasLinhasConsolidado) return;
    await this.baixarDocx(`relatorio_consolidado_${this.competencia()}.docx`, this._ultimasLinhasConsolidado);
  }

  competencia(): string {
    return `${this.ano}-${String(this.mes).padStart(2, "0")}`;
  }
}
