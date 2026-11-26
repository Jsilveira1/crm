"use strict";
/* =====================================================================
   Sistema de Propostas SKA — versão web (Supabase)
   Geração do .docx no navegador com o modelo guardado no banco.
   ===================================================================== */
const VERSAO_APP = "web 1.0 · 26/11/2026";
const sb = supabase.createClient(SKA_CONFIG.supabaseUrl, SKA_CONFIG.supabaseKey, { auth: { persistSession: true, autoRefreshToken: true } });

/* ============================ utilidades ============================ */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const uuid = () => (crypto.randomUUID ? crypto.randomUUID() : "10000000-1000-4000-8000-100000000000".replace(/[018]/g, c => (c ^ crypto.getRandomValues(new Uint8Array(1))[0] & 15 >> c / 4).toString(16)));
const uid = uuid;
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const clone = o => JSON.parse(JSON.stringify(o));
const MESES = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
const fmtBRL = v => "R$ " + (Math.round((+v || 0) * 100) / 100).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtPct = v => (+v || 0).toLocaleString("pt-BR", { maximumFractionDigits: 2 }) + "%";
const hojeISO = () => { const d = new Date(); return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); };
const addDias = (iso, n) => { const d = new Date(iso + "T12:00:00"); d.setDate(d.getDate() + (+n || 0)); return d.toISOString().slice(0, 10); };
const difDias = (a, b) => Math.round((new Date(a + "T12:00:00") - new Date(b + "T12:00:00")) / 864e5);
const dataBR = iso => { if (!iso) return ""; const [y, m, d] = String(iso).slice(0, 10).split("-"); return `${d}/${m}/${y}`; };
const dataExtenso = iso => { if (!iso) return ""; const [y, m, d] = iso.split("-").map(Number); return `${d} de ${MESES[m - 1]} de ${y}`; };
const soDig = s => String(s || "").replace(/\D/g, "");
const fmtCNPJ = s => { const d = soDig(s); return d.length === 14 ? d.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, "$1.$2.$3/$4-$5") : d.length === 11 ? d.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, "$1.$2.$3-$4") : s; };
const fmtFone = s => { const d = soDig(s); if (d.length === 11) return d.replace(/^(\d{2})(\d{5})(\d{4})$/, "($1) $2-$3"); if (d.length === 10) return d.replace(/^(\d{2})(\d{4})(\d{4})$/, "($1) $2-$3"); return s || ""; };
const norm = s => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
function titleCase(s) {
  const minus = new Set(["de", "da", "do", "das", "dos", "e", "em", "a", "o", "para", "com"]);
  const maius = new Set(["LTDA", "ME", "EPP", "EIRELI", "S/A", "SA", "S.A.", "S.A", "CNPJ", "PR", "SC", "SP", "RS", "RJ", "MG", "BR", "II", "III", "IV", "CEP"]);
  return String(s || "").toLowerCase().split(/(\s+)/).map((w, i) => {
    if (/^\s+$/.test(w) || !w) return w;
    const up = w.toUpperCase();
    if (maius.has(up)) return up === "LTDA" ? "Ltda" : up;
    if (i > 0 && minus.has(w)) return w;
    return w.replace(/(^|[-/.(])([a-zà-ú])/g, (m, a, b) => a + b.toUpperCase());
  }).join("");
}
function toast(msg, erro) { const t = $("#toast"); t.textContent = msg; t.className = "on" + (erro ? " erro" : ""); clearTimeout(t._h); t._h = setTimeout(() => t.className = "", erro ? 6000 : 2600); }
function b64ToBuf(b64) { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u.buffer; }
function bufToB64(buf) { const u = new Uint8Array(buf); let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }
function baixarBlob(blob, nome) { const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = nome; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500); }
function nomeArquivo(s) { return String(s).replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, " ").trim(); }
function carregando(on, txt) { $("#carregando").classList.toggle("on", !!on); if (txt) $("#carregandoTxt").textContent = txt; }
const vazioNull = v => (v === undefined || v === null || String(v).trim() === "") ? null : (typeof v === "string" ? v.trim() : v);

/* ============================ banco (Supabase) ============================ */
let usuario = null;
const db = { config: {}, grupos: [], produtos: [], descritivos: [], clientes: [], contatos: [], enderecos: [], telefones: [], licencas: [], propostas: [], modelo: null, modeloInfo: null };

const CONFIG_PADRAO = {
  vendedorNome: "", vendedorFone: "", vendedorEmail: "", cidadeEmissao: "Curitiba/PR", filialAceite: "SKA Curitiba/PR", validadeDias: 22,
  tituloCapa: "Ampliação tecnológica da empresa {empresa}", formaPagamento: "Boleto bancário com a primeira para 15DDF, sem juros.", prazoEntrega: "10 (dez) dias úteis",
  textoCortesia: "R$ Cortesia", descontoPadrao: 20, parcelasPadrao: 24, arredondamento: "cima", reqPadrao: true, matrizPadrao: true,
};
const GRUPOS_PADRAO = ["Software - Licenciamento Perpetuo com 2 anos de Subscription", "Software - Licenciamento Por Locação Anual"];

function erroSb(error, oque) {
  if (!error) return;
  console.error(oque, error);
  let m = error.message || String(error);
  if (/JWT|session|token/i.test(m)) m = "sua sessão expirou — entre novamente.";
  else if (/foreign key/i.test(m)) m = "o registro está sendo usado em outro cadastro (ex.: proposta ou licença).";
  else if (/duplicate key/i.test(m)) m = "já existe um registro com esse valor (código ou título repetido).";
  else if (/Failed to fetch|NetworkError/i.test(m)) m = "sem conexão com a internet.";
  throw new Error("Erro ao " + oque + ": " + m);
}
async function lerTudo(tabela, select = "*", ordem) {
  const lote = 1000; let de = 0, out = [];
  for (;;) {
    let q = sb.from(tabela).select(select).range(de, de + lote - 1);
    if (ordem) q = q.order(ordem);
    const { data, error } = await q; erroSb(error, "ler " + tabela);
    out = out.concat(data); if (data.length < lote) break; de += lote;
  }
  return out;
}
async function gravar(tabela, row) {
  const { data, error } = await sb.from(tabela).upsert(row).select().single();
  erroSb(error, "gravar em " + tabela); return data;
}
async function apagar(tabela, id) { const { error } = await sb.from(tabela).delete().eq("id", id); erroSb(error, "excluir de " + tabela); }
async function gravarConfig(chave, valor) { const { error } = await sb.from("configuracoes").upsert({ chave, valor }); erroSb(error, "gravar configuração"); }
async function tentar(f) { try { return await f(); } catch (e) { toast(e.message, true); return undefined; } }

/* ---------- produtos: banco <-> app ---------- */
function prodDeRow(r) {
  return {
    id: r.id, codigo: r.codigo_crm != null ? String(r.codigo_crm) : "", nome: r.descricao || "", valor: +r.valor || 0,
    negocio: r.negocio || "", linha: r.grupo || "", subgrupo: r.subgrupo || "", tipo: r.tipo_gerador || "",
    anos: r.anos != null ? String(r.anos) : "", meses: r.meses != null ? String(r.meses) : "", licenca: r.local_network || "",
    grupo: r.titulo_tabela || "", cortesia: !!r.cortesia_padrao, familia: r.familia || "", base: r.produto_base || "", conferir: r.conferir || "",
    ativo: r.ativo !== false, descritivos: [],
  };
}
function prodParaRow(p) {
  return {
    id: p.id, codigo_crm: soDig(p.codigo) ? +soDig(p.codigo) : null, descricao: p.nome, valor: +p.valor || 0,
    negocio: vazioNull(p.negocio), grupo: vazioNull(p.linha), subgrupo: vazioNull(p.subgrupo), tipo_gerador: vazioNull(p.tipo),
    anos: p.anos ? +p.anos : null, meses: p.meses ? +p.meses : null, local_network: vazioNull(p.licenca), titulo_tabela: vazioNull(p.grupo),
    cortesia_padrao: !!p.cortesia, familia: vazioNull(p.familia), produto_base: vazioNull(p.base), conferir: vazioNull(p.conferir), ativo: p.ativo !== false,
  };
}
async function carregarTudo() {
  carregando(true, "Carregando dados...");
  try {
    const [cfg, prods, descs, vinc, clis, cts, ends, tels, lics, props] = await Promise.all([
      lerTudo("configuracoes", "chave,valor,atualizado_em"), lerTudo("produtos", "*", "descricao"), lerTudo("descritivos", "*", "titulo"),
      lerTudo("produto_descritivos", "*"), lerTudo("clientes", "*", "nome_fantasia"), lerTudo("contatos", "*", "nome"),
      lerTudo("enderecos", "*"), lerTudo("telefones", "*"), lerTudo("licencas", "*", "validade"), lerTudo("propostas", "*", "atualizado_em"),
    ]);
    const c = {}; for (const r of cfg) if (r.chave !== "modelo_docx") c[r.chave] = r.valor;
    db.config = Object.assign({}, CONFIG_PADRAO, c.proposta_padrao || {});
    db.retroatividade = c.retroatividade || null;
    db.grupos = (c.grupos_tabela && c.grupos_tabela.lista) || [...new Set([...GRUPOS_PADRAO, ...prods.map(p => p.titulo_tabela).filter(Boolean)])];
    db.descritivos = descs.map(d => ({ id: d.id, titulo: d.titulo, texto: d.texto || "" }));
    db.produtos = prods.map(prodDeRow);
    const porId = new Map(db.produtos.map(p => [p.id, p]));
    for (const v of vinc.sort((a, b) => (a.ordem || 0) - (b.ordem || 0))) { const p = porId.get(v.produto_id); if (p) p.descritivos.push(v.descritivo_id); }
    db.clientes = clis; db.contatos = cts; db.enderecos = ends; db.telefones = tels; db.licencas = lics;
    db.propostas = props.reverse();
    const mod = cfg.find(r => r.chave === "modelo_docx");
    db.modeloInfo = mod ? { nome: (mod.valor && mod.valor.nome) || "modelo.docx", em: mod.atualizado_em } : null;
    db.modelo = mod && mod.valor && mod.valor.b64 ? mod.valor.b64 : null;
    db.carregadoEm = new Date();
  } finally { carregando(false); }
}

/* ============================ classificação de produtos ============================ */
const ORDEM_TIPOS = ["Pacote (licença + subscription)", "Licença perpétua (PLC)", "Subscription inicial", "Renovação de subscription (ALC)", "Renovação mensal (MLC)",
  "Locação anual (YLC)", "Locação trimestral (QLC)", "Upgrade / Migração", "Kit / Bundle (composição)", "Taxa / adicional", "Serviço",
  "Capacitação EAD (individual)", "Habilitação (turma fechada)", "Material didático (apostila)"];
const TIPO_CURTO = { "Pacote (licença + subscription)": "Pacote PLC+ALC", "Licença perpétua (PLC)": "Licença PLC", "Subscription inicial": "Sub. inicial", "Renovação de subscription (ALC)": "Renovação ALC",
  "Renovação mensal (MLC)": "Renov. mensal MLC", "Locação anual (YLC)": "Locação YLC", "Locação trimestral (QLC)": "Locação QLC", "Kit / Bundle (composição)": "Kit/Bundle" };
const LICENCAS = ["Local", "Network"];
const CAMPOS_FILTRO = [["negocio", "Negócio"], ["linha", "Grupo"], ["subgrupo", "Subgrupo"], ["tipo", "Tipo"], ["anos", "Anos"], ["licenca", "Licença"]];
const tiposExistentes = () => [...new Set([...ORDEM_TIPOS, ...db.produtos.map(p => p.tipo).filter(Boolean)])];
function classificarPeloNome(nome) {
  const n = nome || "";
  const tipo = /capacita|treinamento/i.test(n) ? "Capacitação EAD (individual)"
    : /\bMLC\b/i.test(n) ? "Renovação mensal (MLC)"
    : /renewal|renova|\bALC\b(?!.*PLC)/i.test(n) && !/PLC\s*\+/i.test(n) ? "Renovação de subscription (ALC)"
    : /PLC\s*\+\s*ALC/i.test(n) ? "Pacote (licença + subscription)"
    : /initial|inicial/i.test(n) ? "Subscription inicial"
    : /\bQLC\b/i.test(n) ? "Locação trimestral (QLC)"
    : /\bYLC\b|term license|loca[cç][aã]o/i.test(n) ? "Locação anual (YLC)"
    : /upgrade|migra/i.test(n) ? "Upgrade / Migração"
    : /[-–]\s*PLC\b|\bPLC$/i.test(n) ? "Licença perpétua (PLC)"
    : /servi[cç]o|implanta|consultoria|\bhoras?\b/i.test(n) ? "Serviço" : "";
  const m = n.match(/(\d+)\s*(years?|anos?)/i);
  const anos = m ? m[1] : /YLC/.test(tipo) ? "1" : "";
  const licenca = /Serviço|Capacita/.test(tipo) ? "" : /network|\brede\b|\bSNL\b/i.test(n) ? "Network" : "Local";
  return { tipo, anos, licenca };
}
function tituloPadraoTabela(p) {
  const a = +p.anos || 0, t = p.tipo || "";
  if (t === "Pacote (licença + subscription)") return `Software - Licenciamento Perpetuo com ${a || 2} anos de Subscription`;
  if (t === "Licença perpétua (PLC)") return "Software - Licenciamento Perpetuo";
  if (t === "Locação anual (YLC)") return "Software - Licenciamento Por Locação Anual";
  if (t === "Locação trimestral (QLC)") return "Software - Licenciamento Por Locação Trimestral";
  if (t === "Renovação mensal (MLC)") return "Software - Renovação Mensal de Subscription";
  if (t === "Renovação de subscription (ALC)") return `Software - Renovação de Subscription${a ? ` (${a} ano${a > 1 ? "s" : ""})` : ""}`;
  if (t === "Subscription inicial") return `Software - Subscription${a ? ` (${a} ano${a > 1 ? "s" : ""})` : ""}`;
  if (t === "Upgrade / Migração") return "Software - Upgrade / Migração";
  if (/Capacita|Habilita|Material did/.test(t)) return "Treinamentos";
  if (t === "Serviço") return "Serviços";
  return "";
}

/* ============================ clientes: auxiliares ============================ */
const contatosDe = id => db.contatos.filter(x => x.cliente_id === id);
const enderecosDe = id => db.enderecos.filter(x => x.cliente_id === id).sort((a, b) => (b.principal ? 1 : 0) - (a.principal ? 1 : 0));
const telefonesDe = id => db.telefones.filter(x => x.cliente_id === id);
const licencasDe = id => db.licencas.filter(x => x.cliente_id === id);
const cliPorId = id => db.clientes.find(c => c.id === id);
const empresaDe = c => c.razao_social || c.nome_fantasia || "";
function rotuloCliente(c) {
  const r = c.razao_social && norm(c.razao_social) !== norm(c.nome_fantasia) ? " · " + c.razao_social : "";
  return (c.nome_fantasia || c.razao_social || "") + r + (c.cnpj_cpf ? " — " + c.cnpj_cpf : "");
}
function textoEndereco(e) {
  if (!e) return "";
  return [[e.logradouro, e.numero].filter(Boolean).join(", "), e.complemento, e.bairro].filter(Boolean).join(" - ");
}
function textoTelefone(t) { if (!t) return ""; const n = fmtFone((t.ddd || "") + (t.numero || "")); return n + (t.ramal ? " r." + t.ramal : ""); }
function corSuporte(s) { s = norm(s); return /verde/.test(s) ? "b-verde" : /amarelo/.test(s) ? "b-amarelo" : /vermelho/.test(s) ? "b-vermelho" : "b-cinza"; }
function badgeSuporte(s) { return s ? `<span class="badge ${corSuporte(s)}">${esc(String(s).replace(/\s*\(.*\)/, ""))}</span>` : ""; }
function badgeNivel(n) { return n ? `<span class="badge ${/ouro/i.test(n) ? "b-amarelo" : /prata/i.test(n) ? "b-cinza" : "b-azul"}">${esc(n)}</span>` : ""; }
function situacaoLicenca(l) {
  if (!l.validade) return { cod: "sem", txt: "Sem validade", cls: "b-cinza", dias: null, meses: 0 };
  const dias = difDias(l.validade, hojeISO());
  if (dias >= 0) return { cod: dias <= 90 ? "avencer" : "ok", txt: dias === 0 ? "Vence hoje" : `Vence em ${dias} dia${dias > 1 ? "s" : ""}`, cls: dias <= 30 ? "b-vermelho" : dias <= 90 ? "b-amarelo" : "b-verde", dias, meses: 0 };
  const meses = mesesVencidos(l.validade);
  return { cod: "vencida", txt: `Vencida há ${meses} ${meses > 1 ? "meses" : "mês"}`, cls: "b-vermelho", dias, meses };
}
function mesesVencidos(validade) { // igual à view vw_licencas_situacao: meses completos + 1 se sobrar algum dia
  const v = new Date(validade + "T12:00:00"), h = new Date(hojeISO() + "T12:00:00");
  let m = (h.getFullYear() - v.getFullYear()) * 12 + (h.getMonth() - v.getMonth());
  if (h.getDate() < v.getDate()) m--;
  const ref = new Date(v); ref.setMonth(ref.getMonth() + m);
  return m + (h > ref ? 1 : 0);
}

/* ============================ proposta atual ============================ */
const ORIGENS = ["Prospecção", "Renovação", "Vencida", "Compliance", "Ampliação", "Outra"];
const SITUACOES = ["Rascunho", "Enviada", "Em negociação", "Ganha", "Perdida", "Cancelada"];
const LS_RASCUNHO = "ska-web-rascunho";
let cur = null;
const clienteVazio = () => ({ id: null, contatoId: null, empresa: "", cnpj: "", contato: "", email: "", telefone: "", endereco: "", cidade: "", uf: "" });
function novaProposta() {
  const c = db.config;
  return {
    id: uuid(), numero: "", data: hojeISO(), validade: addDias(hojeISO(), c.validadeDias), tituloCapa: "", tituloManual: false,
    cidadeEmissao: c.cidadeEmissao, cliente: clienteVazio(), origem: "Prospecção", situacao: "Rascunho",
    modo: "opcoes", parcelasGeral: c.parcelasPadrao, tabelas: [novaTabela()],
    descExcl: [], descAdd: [], descOrdem: [],
    formaPagamento: c.formaPagamento, prazoEntrega: c.prazoEntrega, mostrarRequisitos: !!c.reqPadrao, mostrarMatriz: !!c.matrizPadrao,
  };
}
function novaTabela(titulo) {
  return { id: uid(), titulo: titulo || (db.grupos[0] || ""), itens: [], descPct: db.config.descontoPadrao, parcelas: db.config.parcelasPadrao, rotulo: "" };
}
function salvarRascunho() { try { localStorage.setItem(LS_RASCUNHO, JSON.stringify(cur)); } catch { } }

function calcTabela(t) {
  let sub = 0;
  for (const it of t.itens) if (!it.cortesia) sub += (+it.qtd || 0) * (+it.valor || 0);
  sub = Math.round(sub * 100) / 100;
  const desc = Math.round(sub * (+t.descPct || 0)) / 100;
  const total = Math.round((sub - desc) * 100) / 100;
  const n = Math.max(1, parseInt(t.parcelas) || 1);
  return { sub, desc, total, n, parcela: parcelaDe(total, n) };
}
function parcelaDe(total, n) {
  const raw = total / n;
  return db.config.arredondamento === "normal" ? Math.round(raw * 100) / 100 : Math.ceil(Math.round(raw * 1e6) / 1e4) / 100;
}
function tituloCapa() {
  if (cur.tituloManual && cur.tituloCapa) return cur.tituloCapa;
  return (db.config.tituloCapa || "").replace(/\{empresa\}/g, cur.cliente.empresa || "");
}
function rotuloTabela(t, i) { if (t.rotulo) return t.rotulo; return cur.modo === "opcoes" ? `(    ) Opção ${i + 1}` : ""; }

/* -------- descritivos derivados -------- */
function descAuto() {
  const ids = [], maxLinhas = Math.max(0, ...cur.tabelas.map(t => t.itens.length));
  for (let r = 0; r < maxLinhas; r++) for (const t of cur.tabelas) {
    const it = t.itens[r]; if (!it) continue;
    const p = db.produtos.find(p => p.id === it.prodId) || db.produtos.find(p => norm(p.nome) === norm(it.nome));
    if (p) for (const d of p.descritivos || []) if (!ids.includes(d) && db.descritivos.some(x => x.id === d)) ids.push(d);
  }
  return ids;
}
function descLista() {
  const auto = descAuto();
  let todos = [...auto, ...cur.descAdd.filter(i => !auto.includes(i) && db.descritivos.some(x => x.id === i))];
  const ord = cur.descOrdem || [];
  todos.sort((a, b) => { const ia = ord.indexOf(a), ib = ord.indexOf(b); return (ia < 0 ? 1e6 + todos.indexOf(a) : ia) - (ib < 0 ? 1e6 + todos.indexOf(b) : ib); });
  return todos.map(id => ({ id, auto: auto.includes(id), on: !cur.descExcl.includes(id) }));
}

/* ============================ render: proposta ============================ */
function renderProposta() {
  for (const el of $$("[data-c]")) el.value = cur.cliente[el.dataset.c] || "";
  const c = cur.cliente.id && cliPorId(cur.cliente.id);
  $("#pCliInfo").textContent = c ? "Vinculado ao cadastro: " + (c.nome_fantasia || c.razao_social) + (c.codigo_crm ? " (CRM " + c.codigo_crm + ")" : "") : "";
  $("#pBtnVerCli").style.display = c ? "" : "none";
  renderAlertaCliente(c);
  const cts = c ? contatosDe(c.id).filter(x => x.ativo !== false) : [];
  $("#dlContatos").innerHTML = cts.map(x => `<option value="${esc(x.nome)}">${esc([x.decisor ? "★ decisor" : "", x.cargo, x.email].filter(Boolean).join(" · "))}</option>`).join("");
  $("#pContato").placeholder = cts.length ? `Digite ou escolha um dos ${cts.length} contato(s) do cliente` : "Nome de quem pediu a proposta";
  $("#pContatoInfo").textContent = !c ? "" : cts.length ? `${cts.length} contato(s) cadastrado(s) — apague o campo para ver a lista.` : "Cliente sem contatos cadastrados: digite o nome e use \"Salvar / atualizar no cadastro\" para gravá-lo.";
  $("#pNumero").value = cur.numero; $("#pData").value = cur.data; $("#pValidade").value = cur.validade;
  $("#pTitulo").value = tituloCapa(); $("#pCidadeEmissao").value = cur.cidadeEmissao;
  $("#pOrigem").innerHTML = ORIGENS.map(o => `<option ${o === cur.origem ? "selected" : ""}>${o}</option>`).join("");
  $("#pSituacao").innerHTML = SITUACOES.map(o => `<option ${o === cur.situacao ? "selected" : ""}>${o}</option>`).join("");
  $$("input[name=modo]").forEach(r => r.checked = r.value === cur.modo);
  $("#pFormaPag").value = cur.formaPagamento; $("#pPrazo").value = cur.prazoEntrega;
  $("#pReq").checked = cur.mostrarRequisitos; $("#pMatriz").checked = cur.mostrarMatriz;
  $("#pParcGeral").value = cur.parcelasGeral;
  renderTabelas(); renderDescs(); renderResumo();
}
function renderAlertaCliente(c) {
  const box = $("#pCliAlerta"); if (!c) { box.innerHTML = ""; return; }
  const av = [];
  if (c.suporte && !/verde/i.test(c.suporte)) av.push("Suporte: " + c.suporte);
  if (c.status_licenciamento && /inativo/i.test(c.status_licenciamento)) av.push("Licenciamento inativo");
  const venc = licencasDe(c.id).map(situacaoLicenca).filter(s => s.cod === "vencida");
  if (venc.length) av.push(`${venc.length} licença(s) vencida(s) — maior atraso: ${Math.max(...venc.map(s => s.meses))} mês(es)`);
  if (c.pontos_atencao) av.push(c.pontos_atencao);
  box.innerHTML = av.length ? `<div class="alerta-cli aviso" style="display:block"><b>Atenção:</b> ${av.map(esc).join(" · ")}</div>` : "";
}
function renderTabelas() {
  const box = $("#pTabelas");
  box.innerHTML = cur.tabelas.map((t, ti) => `
  <div class="tabela-prop" data-t="${ti}">
    <div class="topo">
      <span class="pill">Tabela ${ti + 1}</span>
      <input class="cresce" data-tf="titulo" list="dlGrupos" value="${esc(t.titulo)}" placeholder="Título da tabela" style="flex:1;min-width:0">
      <button class="btn peq" data-ta="up" title="Mover para cima">▲</button>
      <button class="btn peq" data-ta="down" title="Mover para baixo">▼</button>
      <button class="btn peq perigo" data-ta="del" title="Excluir tabela">✕</button>
    </div>
    <div class="corpo tab-itens">
      <table class="lista">
        <thead><tr><th style="width:70px">Qtde</th><th>Item</th><th style="width:140px" class="num-r">Valor unitário</th><th style="width:70px">Cortesia</th><th style="width:130px" class="num-r">Valor total</th><th style="width:80px"></th></tr></thead>
        <tbody>
        ${t.itens.map((it, ii) => `<tr data-i="${ii}">
          <td><input type="number" min="0" step="1" data-if="qtd" value="${esc(it.qtd)}"></td>
          <td><input data-if="nome" value="${esc(it.nome)}"></td>
          <td><input type="number" min="0" step="0.01" data-if="valor" value="${esc(it.valor)}" class="num-r"></td>
          <td style="text-align:center"><input type="checkbox" data-if="cortesia" ${it.cortesia ? "checked" : ""}></td>
          <td class="num-r">${it.cortesia ? esc(db.config.textoCortesia) : fmtBRL((+it.qtd || 0) * (+it.valor || 0))}</td>
          <td style="white-space:nowrap"><button class="btn peq" data-ia="up">▲</button><button class="btn peq perigo" data-ia="del">✕</button></td>
        </tr>`).join("") || `<tr><td colspan="6" class="vazio">Nenhum item. Adicione abaixo.</td></tr>`}
        </tbody>
      </table>
    </div>
    <div class="corpo" style="padding-top:0">
      <div class="linha" style="margin-top:4px">
        <div class="cresce ac"><input data-add autocomplete="off" placeholder="Adicionar produto: digite partes do nome (ex.: prem anual)..."><div class="ac-lista"></div></div>
        <button class="btn" data-ta="catalogo">Catálogo…</button>
        <button class="btn" data-ta="additem" title="Item que não está no cadastro">+ Avulso</button>
      </div>
      ${(() => { const c = calcTabela(t); return `
      <div class="totais">
        <span>Desconto (%)</span><input type="number" step="0.01" min="0" data-tf="descPct" value="${esc(t.descPct)}">
        <span>Parcelas</span><input type="number" min="1" data-tf="parcelas" value="${esc(t.parcelas)}">
        <span>Rótulo</span><input data-tf="rotulo" value="${esc(t.rotulo)}" placeholder="${esc(rotuloTabela({ rotulo: "" }, ti)) || "(sem rótulo)"}" style="width:150px">
      </div>
      <div class="totais" style="grid-template-columns:repeat(8,auto)">
        <span>SubTotal</span><b>${fmtBRL(c.sub)}</b>
        <span>Desconto ${fmtPct(t.descPct)}</span><b>${fmtBRL(c.desc)}</b>
        <span>Total</span><b style="color:var(--azul-esc)">${fmtBRL(c.total)}</b>
        <span>Simulado em ${c.n}x</span><b>${fmtBRL(c.parcela)}</b>
      </div>`; })()}
    </div>
  </div>`).join("") || `<div class="vazio">Nenhuma tabela.</div>`;
  const soma = cur.modo === "somatoria";
  $("#pSomaBox").style.display = soma ? "" : "none";
  if (soma) { const tg = totalGeral(); $("#pTotalGeral").textContent = fmtBRL(tg); $("#pParcGeralVal").textContent = fmtBRL(parcelaDe(tg, Math.max(1, +cur.parcelasGeral || 1))); }
}
function totalGeral() { return Math.round(cur.tabelas.reduce((s, t) => s + calcTabela(t).total, 0) * 100) / 100; }
function renderDescs() {
  const l = descLista();
  $("#pDescs").innerHTML = l.map((d, i) => { const x = db.descritivos.find(z => z.id === d.id); return `
    <div class="desc-item ${d.on ? "" : "off"}" data-d="${d.id}">
      <input type="checkbox" ${d.on ? "checked" : ""} data-da="tog">
      <span class="t">${esc(x.titulo)} ${d.auto ? `<span class="pill">do produto</span>` : `<span class="pill" style="background:#eef7ee;color:#1e7b46">incluído manualmente</span>`}</span>
      <button class="btn peq" data-da="ver">Ver</button>
      <button class="btn peq" data-da="up" ${i === 0 ? "disabled" : ""}>▲</button>
      <button class="btn peq" data-da="down" ${i === l.length - 1 ? "disabled" : ""}>▼</button>
      ${d.auto ? "" : `<button class="btn peq perigo" data-da="rm">✕</button>`}
    </div>`; }).join("") || `<div class="vazio">Nenhum descritivo. Adicione produtos às tabelas ou inclua um descritivo abaixo.</div>`;
  const usados = new Set(l.map(d => d.id));
  $("#pDescAdd").innerHTML = `<option value="">Incluir outro descritivo...</option>` + db.descritivos.filter(d => !usados.has(d.id)).map(d => `<option value="${d.id}">${esc(d.titulo)}</option>`).join("");
}
function renderResumo() {
  const n = cur.tabelas.length;
  const tot = cur.modo === "somatoria" ? "Total geral " + fmtBRL(totalGeral()) : cur.tabelas.map((t, i) => `Op.${i + 1}: ${fmtBRL(calcTabela(t).total)}`).join("  ·  ");
  $("#resumo").textContent = `${cur.numero ? "Proposta " + cur.numero + " — " : ""}${cur.cliente.empresa || "sem cliente"} — ${n} tabela(s)${n ? " — " + tot : ""}`;
}
function mudou(reRender) { salvarRascunho(); if (reRender === "tab") { renderTabelas(); renderDescs(); } renderResumo(); }

/* -------- aplicar cliente do cadastro na proposta -------- */
function aplicarCliente(c, contatoId) {
  const cts = contatosDe(c.id).filter(x => x.ativo !== false);
  const ct = cts.find(x => x.id === contatoId) || cts.find(x => x.decisor) || cts[0];
  const end = enderecosDe(c.id)[0], tel = telefonesDe(c.id)[0];
  cur.cliente = {
    id: c.id, contatoId: ct ? ct.id : null, empresa: empresaDe(c), cnpj: c.cnpj_cpf || "",
    contato: ct ? ct.nome : "", email: ct ? ct.email || "" : "", telefone: (ct && (ct.telefone || ct.celular)) || textoTelefone(tel),
    endereco: textoEndereco(end), cidade: (end && end.cidade) || c.cidade || "", uf: (end && end.uf) || c.uf || "",
  };
  if (!cur.tituloManual) cur.tituloCapa = "";
}
function acharClientePorTexto(txt) {
  const v = norm(txt).trim(); if (!v) return null;
  return db.clientes.find(c => norm(rotuloCliente(c)) === v)
    || (soDig(v).length >= 8 ? db.clientes.find(c => soDig(c.cnpj_cpf).startsWith(soDig(v))) : null)
    || db.clientes.find(c => norm(c.nome_fantasia) === v || norm(c.razao_social) === v)
    || db.clientes.find(c => norm(c.nome_fantasia + " " + c.razao_social).includes(v));
}

/* -------- eventos proposta -------- */
function ligarProposta() {
  $$("[data-c]").forEach(el => el.addEventListener("input", () => {
    cur.cliente[el.dataset.c] = el.value;
    if (el.dataset.c === "empresa" && !cur.tituloManual) $("#pTitulo").value = tituloCapa();
    mudou();
  }));
  $("#pCliBusca").addEventListener("change", e => {
    const c = acharClientePorTexto(e.target.value); if (!e.target.value.trim()) return;
    if (!c) return toast("Cliente não encontrado no cadastro.", true);
    aplicarCliente(c); e.target.value = ""; renderProposta(); salvarRascunho(); toast("Cliente carregado: " + (c.nome_fantasia || c.razao_social));
  });
  $("#pContato").addEventListener("change", e => {
    const v = norm(e.target.value).trim(); if (!v || !cur.cliente.id) return;
    const ct = contatosDe(cur.cliente.id).find(x => norm(x.nome) === v); if (!ct) { cur.cliente.contatoId = null; return salvarRascunho(); }
    Object.assign(cur.cliente, { contatoId: ct.id, contato: ct.nome, email: ct.email || cur.cliente.email, telefone: ct.telefone || ct.celular || cur.cliente.telefone });
    renderProposta(); salvarRascunho();
  });
  $("#pBtnVerCli").addEventListener("click", () => { const c = cliPorId(cur.cliente.id); if (c) modalCliente(c); });
  $("#pBtnCnpj").addEventListener("click", async () => {
    const btn = $("#pBtnCnpj"), cnpj = $("#pCnpjBusca").value || cur.cliente.cnpj;
    btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Buscando';
    try {
      const r = await buscarCNPJ(cnpj);
      const existente = db.clientes.find(c => soDig(c.cnpj_cpf) && soDig(c.cnpj_cpf) === soDig(r.cnpj));
      if (existente) { aplicarCliente(existente); for (const k of ["empresa", "endereco", "cidade", "uf"]) if (r[k]) cur.cliente[k] = r[k]; if (!cur.cliente.telefone) cur.cliente.telefone = r.telefone; }
      else cur.cliente = Object.assign(clienteVazio(), r);
      renderProposta(); salvarRascunho();
      toast(existente ? "Cliente já cadastrado — dados atualizados pela Receita (salve para gravar)." : "Dados carregados da Receita Federal. Confira e salve no cadastro.");
    } catch (e) { toast(e.message, true); }
    finally { btn.disabled = false; btn.textContent = "Buscar CNPJ"; }
  });
  $("#pCnpjBusca").addEventListener("keydown", e => { if (e.key === "Enter") $("#pBtnCnpj").click(); });
  $("#pBtnSalvarCli").addEventListener("click", () => tentar(salvarClienteDaProposta));
  $("#pNumero").addEventListener("input", e => { cur.numero = e.target.value.trim(); mudou(); });
  $("#pData").addEventListener("change", e => { cur.data = e.target.value; mudou(); });
  $("#pValidade").addEventListener("change", e => { cur.validade = e.target.value; mudou(); });
  $$("[data-val]").forEach(b => b.addEventListener("click", () => { cur.validade = addDias(cur.data || hojeISO(), b.dataset.val); $("#pValidade").value = cur.validade; mudou(); }));
  $("#pTitulo").addEventListener("input", e => { cur.tituloCapa = e.target.value; cur.tituloManual = e.target.value !== (db.config.tituloCapa || "").replace(/\{empresa\}/g, cur.cliente.empresa || ""); mudou(); });
  $("#pCidadeEmissao").addEventListener("input", e => { cur.cidadeEmissao = e.target.value; mudou(); });
  $("#pOrigem").addEventListener("change", e => { cur.origem = e.target.value; mudou(); });
  $("#pSituacao").addEventListener("change", e => { cur.situacao = e.target.value; mudou(); });
  $$("input[name=modo]").forEach(r => r.addEventListener("change", () => { cur.modo = r.value; mudou("tab"); }));
  $("#pParcGeral").addEventListener("input", e => { cur.parcelasGeral = e.target.value; renderTabelas(); mudou(); });
  $("#pFormaPag").addEventListener("input", e => { cur.formaPagamento = e.target.value; mudou(); });
  $("#pPrazo").addEventListener("input", e => { cur.prazoEntrega = e.target.value; mudou(); });
  $("#pReq").addEventListener("change", e => { cur.mostrarRequisitos = e.target.checked; mudou(); });
  $("#pMatriz").addEventListener("change", e => { cur.mostrarMatriz = e.target.checked; mudou(); });
  $("#pBtnAddTab").addEventListener("click", () => { cur.tabelas.push(novaTabela()); mudou("tab"); });
  $("#pBtnDupTab").addEventListener("click", () => { const u = cur.tabelas[cur.tabelas.length - 1]; const n = u ? Object.assign(clone(u), { id: uid(), rotulo: "" }) : novaTabela(); cur.tabelas.push(n); mudou("tab"); toast("Tabela duplicada — troque o produto principal."); });

  const box = $("#pTabelas");
  box.addEventListener("change", e => {
    const tEl = e.target.closest("[data-t]"); if (!tEl) return; const t = cur.tabelas[+tEl.dataset.t];
    if (e.target.dataset.tf) { t[e.target.dataset.tf] = e.target.value; mudou("tab"); }
    else if (e.target.dataset.if) {
      const it = t.itens[+e.target.closest("[data-i]").dataset.i], f = e.target.dataset.if;
      if (f === "cortesia") it.cortesia = e.target.checked;
      else if (f === "nome") { it.nome = e.target.value; const p = acharProduto(it.nome); if (p) { it.prodId = p.id; it.nome = p.nome; it.valor = p.valor; it.cortesia = !!p.cortesia; } else it.prodId = null; }
      else it[f] = e.target.value;
      mudou("tab");
    }
  });
  box.addEventListener("input", e => { if (e.target.matches("[data-add]")) acMostrar(e.target); });
  box.addEventListener("focusin", e => { if (e.target.matches("[data-add]") && e.target.value.trim()) acMostrar(e.target); });
  box.addEventListener("focusout", e => { if (e.target.matches("[data-add]")) setTimeout(() => { const l = e.target.parentElement && e.target.parentElement.querySelector(".ac-lista"); if (l) l.classList.remove("on"); }, 200); });
  box.addEventListener("mousedown", e => { if (e.target.closest(".ac-op")) e.preventDefault(); }); // mantém o foco no campo
  box.addEventListener("click", e => { // clique (PC) ou toque (celular) numa sugestão
    const op = e.target.closest(".ac-op"); if (!op) return;
    const t = cur.tabelas[+op.closest("[data-t]").dataset.t]; addProduto(t, db.produtos.find(p => p.id === op.dataset.p));
  });
  box.addEventListener("keydown", e => {
    if (!e.target.matches("[data-add]")) return;
    const lista = e.target.parentElement.querySelector(".ac-lista"); const ops = $$(".ac-op", lista); let i = ops.findIndex(o => o.classList.contains("sel"));
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault(); if (!lista.classList.contains("on")) return acMostrar(e.target);
      i = e.key === "ArrowDown" ? Math.min(ops.length - 1, i + 1) : Math.max(0, i - 1);
      ops.forEach((o, k) => o.classList.toggle("sel", k === i)); if (ops[i]) ops[i].scrollIntoView({ block: "nearest" });
    } else if (e.key === "Enter") {
      e.preventDefault(); const t = cur.tabelas[+e.target.closest("[data-t]").dataset.t];
      const op = ops[i >= 0 ? i : 0];
      if (op && lista.classList.contains("on")) addProduto(t, db.produtos.find(p => p.id === op.dataset.p));
      else if (e.target.value.trim()) addItem(t, e.target.value);
    } else if (e.key === "Escape") lista.classList.remove("on");
  });
  box.addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return;
    const tEl = b.closest("[data-t]"); const ti = +tEl.dataset.t; const t = cur.tabelas[ti];
    if (b.dataset.ta === "del") { if (t.itens.length && !confirm("Excluir esta tabela?")) return; cur.tabelas.splice(ti, 1); }
    else if (b.dataset.ta === "up" && ti > 0) [cur.tabelas[ti - 1], cur.tabelas[ti]] = [cur.tabelas[ti], cur.tabelas[ti - 1]];
    else if (b.dataset.ta === "down" && ti < cur.tabelas.length - 1) [cur.tabelas[ti + 1], cur.tabelas[ti]] = [cur.tabelas[ti], cur.tabelas[ti + 1]];
    else if (b.dataset.ta === "additem") { const inp = $("[data-add]", tEl); return addItem(t, inp.value || "", true); }
    else if (b.dataset.ta === "catalogo") return modalCatalogo(t);
    else if (b.dataset.ia) {
      const ii = +b.closest("[data-i]").dataset.i;
      if (b.dataset.ia === "del") t.itens.splice(ii, 1);
      if (b.dataset.ia === "up" && ii > 0) [t.itens[ii - 1], t.itens[ii]] = [t.itens[ii], t.itens[ii - 1]];
    } else return;
    mudou("tab");
  });
  $("#pDescs").addEventListener("click", e => {
    const el = e.target.closest("[data-d]"); if (!el) return; const id = el.dataset.d; const a = e.target.dataset.da; if (!a) return;
    const lista = descLista().map(d => d.id); const i = lista.indexOf(id);
    if (a === "tog") { if (e.target.checked) cur.descExcl = cur.descExcl.filter(x => x !== id); else cur.descExcl.push(id); }
    else if (a === "up" && i > 0) { [lista[i - 1], lista[i]] = [lista[i], lista[i - 1]]; cur.descOrdem = lista; }
    else if (a === "down" && i < lista.length - 1) { [lista[i + 1], lista[i]] = [lista[i], lista[i + 1]]; cur.descOrdem = lista; }
    else if (a === "rm") { cur.descAdd = cur.descAdd.filter(x => x !== id); }
    else if (a === "ver") { const d = db.descritivos.find(x => x.id === id); return abrirModal(`<h3>${esc(d.titulo)}</h3><div class="preview-desc" style="max-height:60vh">${previewTexto(d.texto)}</div><div class="rodape"><button class="btn" onclick="fecharModal()">Fechar</button></div>`); }
    mudou(); renderDescs();
  });
  $("#pBtnDescAdd").addEventListener("click", () => { const v = $("#pDescAdd").value; if (!v) return; cur.descAdd.push(v); cur.descExcl = cur.descExcl.filter(x => x !== v); mudou(); renderDescs(); });
}
async function salvarClienteDaProposta() {
  const k = cur.cliente;
  if (!k.empresa) return toast("Informe ao menos o nome da empresa.", true);
  let c = k.id && cliPorId(k.id);
  if (!c && soDig(k.cnpj)) c = db.clientes.find(x => soDig(x.cnpj_cpf) === soDig(k.cnpj));
  const d = soDig(k.cnpj);
  const base = c ? { ...c } : { id: uuid(), nome_fantasia: k.empresa, ativo: true };
  Object.assign(base, { razao_social: k.empresa, cnpj_cpf: vazioNull(fmtCNPJ(k.cnpj)), cidade: vazioNull(k.cidade), uf: vazioNull((k.uf || "").toUpperCase()),
    tipo_documento: d.length === 14 ? "PJ (CNPJ)" : d.length === 11 ? "PF (CPF)" : (base.tipo_documento || null), cnpj_raiz: d.length === 14 ? fmtCNPJ(d).slice(0, 10) : (base.cnpj_raiz || null) });
  delete base.criado_em; delete base.atualizado_em;
  const salvo = await gravar("clientes", base);
  substituir(db.clientes, salvo);
  // endereço principal
  if (k.endereco) {
    const e = enderecosDe(salvo.id)[0];
    if (!e || textoEndereco(e) !== k.endereco || e.cidade !== k.cidade || e.uf !== k.uf) {
      const row = e ? { ...e, logradouro: k.endereco, numero: null, complemento: null, bairro: null, cidade: vazioNull(k.cidade), uf: vazioNull(k.uf) }
        : { id: uuid(), cliente_id: salvo.id, tipo: "Comercial", logradouro: k.endereco, cidade: vazioNull(k.cidade), uf: vazioNull(k.uf), principal: true };
      substituir(db.enderecos, await gravar("enderecos", row));
    }
  }
  // contato
  if (k.contato) {
    let ct = (k.contatoId && db.contatos.find(x => x.id === k.contatoId)) || contatosDe(salvo.id).find(x => norm(x.nome) === norm(k.contato));
    if (ct && norm(ct.nome) !== norm(k.contato)) ct = null;
    const row = ct ? { ...ct, email: vazioNull(k.email) || ct.email, telefone: vazioNull(k.telefone) || ct.telefone }
      : { id: uuid(), cliente_id: salvo.id, nome: k.contato, email: vazioNull(k.email), telefone: vazioNull(k.telefone) };
    delete row.criado_em;
    const sc = await gravar("contatos", row); substituir(db.contatos, sc); k.contatoId = sc.id;
  }
  k.id = salvo.id; salvarRascunho(); renderDatalists(); renderProposta(); renderListaCli();
  toast(c ? "Cadastro do cliente atualizado." : "Cliente incluído no cadastro.");
}
function substituir(lista, row) { const i = lista.findIndex(x => x.id === row.id); if (i >= 0) lista[i] = row; else lista.push(row); }
function addItem(t, texto, avulso) {
  const p = avulso ? null : acharProduto(texto);
  if (p) return addProduto(t, p);
  t.itens.push({ qtd: 1, prodId: null, nome: texto.trim(), valor: 0, cortesia: false });
  mudou("tab");
  toast("Item avulso (fora do cadastro): preencha o nome e o valor na tabela.");
}
function addProduto(t, p, qtd) {
  if (!p) return;
  t.itens.push({ qtd: qtd || 1, prodId: p.id, nome: p.nome, valor: p.valor, cortesia: !!p.cortesia });
  if (t.itens.length === 1) { const tt = p.grupo || tituloPadraoTabela(p); if (tt) t.titulo = tt; }
  if (p.conferir) toast("Atenção — " + p.nome + ": " + p.conferir);
  mudou("tab");
  const inp = $(`[data-t="${cur.tabelas.indexOf(t)}"] [data-add]`); if (inp && matchMedia("(min-width: 761px)").matches) inp.focus();
}
/* ---------- busca de produtos ---------- */
function textoBuscaProd(p) {
  if (!p._busca) p._busca = norm([p.nome, p.codigo, p.negocio, p.linha, p.subgrupo, p.tipo, TIPO_CURTO[p.tipo], p.anos ? p.anos + " anos " + p.anos + " years" : "", p.licenca, p.grupo, p.base].join(" "));
  return p._busca;
}
function filtrarProdutos(q, f, incluirInativos) {
  f = f || {};
  const termos = norm(q).split(/\s+/).filter(Boolean);
  return db.produtos.filter(p => (incluirInativos || p.ativo) && CAMPOS_FILTRO.every(([k]) => !f[k] || String(p[k] || "") === f[k]) && termos.every(tm => textoBuscaProd(p).includes(tm)))
    .sort((a, b) => (a.negocio || "~").localeCompare(b.negocio || "~") || (a.linha || "~").localeCompare(b.linha || "~") || a.nome.localeCompare(b.nome));
}
function realce(txt, q) {
  const termos = norm(q).split(/\s+/).filter(t => t.length > 1); if (!termos.length) return esc(txt);
  const base = norm(txt); const marcas = new Array(txt.length).fill(false);
  for (const tm of termos) { let i = base.indexOf(tm); while (i >= 0) { for (let k = i; k < i + tm.length; k++) marcas[k] = true; i = base.indexOf(tm, i + 1); } }
  let out = "", aberto = false;
  for (let k = 0; k < txt.length; k++) { if (marcas[k] && !aberto) { out += "<mark>"; aberto = true; } if (!marcas[k] && aberto) { out += "</mark>"; aberto = false; } out += esc(txt[k]); }
  return out + (aberto ? "</mark>" : "");
}
function etiquetas(p) {
  const e = [];
  if (p.tipo) e.push(`<span class="pill">${esc(TIPO_CURTO[p.tipo] || p.tipo)}</span>`);
  if (p.anos) e.push(`<span class="pill" style="background:#eef7ee;color:#1e7b46">${esc(p.anos)} ano${+p.anos > 1 ? "s" : ""}</span>`);
  if (p.meses) e.push(`<span class="pill" style="background:#eef7ee;color:#1e7b46">${esc(p.meses)} ${+p.meses > 1 ? "meses" : "mês"}</span>`);
  if (p.licenca === "Network") e.push(`<span class="pill" style="background:#fdf1e3;color:#9a5b00">Network</span>`);
  if (p.cortesia) e.push(`<span class="pill" style="background:#f1f1f1;color:#555">cortesia</span>`);
  if (!p.ativo) e.push(`<span class="pill" style="background:#fdecea;color:#b42318">inativo</span>`);
  if (p.conferir) e.push(`<span class="pill" style="background:#fff6dd;color:#9a6700" title="${esc(p.conferir)}">conferir</span>`);
  return e.join(" ");
}
function cabecalhoGrupo(p) { return [p.negocio || "Sem negócio", p.linha, p.subgrupo !== p.linha ? p.subgrupo : ""].filter(Boolean).join(" › "); }
function acMostrar(inp) {
  const lista = inp.parentElement.querySelector(".ac-lista"), q = inp.value.trim();
  if (!q) { lista.classList.remove("on"); return; }
  const res = filtrarProdutos(q), max = 15;
  lista.innerHTML = res.slice(0, max).map((p, k) => `<div class="ac-op${k === 0 ? " sel" : ""}" data-p="${p.id}"><span class="n">${realce(p.nome, q)} ${etiquetas(p)}<div class="dica" style="margin:1px 0 0">${esc(cabecalhoGrupo(p))}${p.codigo ? " · Cód. " + esc(p.codigo) : ""}</div></span><span class="v">${fmtBRL(p.valor)}</span></div>`).join("")
    + (res.length > max ? `<div class="ac-rod">Mostrando ${max} de ${res.length}. Digite mais palavras ou use o "Catálogo".</div>` : "")
    + (!res.length ? `<div class="ac-rod">Nenhum produto encontrado. Enter adiciona como item avulso.</div>` : "");
  lista.classList.add("on");
}
/* ---------- filtros em cascata ---------- */
function barraFiltros(pref) { return CAMPOS_FILTRO.map(([k, r]) => `<select id="${pref}-${k}" data-filtro="${k}" title="${r}"></select>`).join(""); }
function atualizarFiltros(pref) {
  const f = {}; const TIPOS = tiposExistentes();
  for (const [k, r] of CAMPOS_FILTRO) {
    const el = $(`#${pref}-${k}`); if (!el) continue;
    const base = db.produtos.filter(p => Object.keys(f).every(c => !f[c] || String(p[c] || "") === f[c]));
    const cont = {}; for (const p of base) { const v = String(p[k] || ""); if (v) cont[v] = (cont[v] || 0) + 1; }
    let vals = Object.keys(cont);
    vals = k === "anos" ? vals.sort((a, b) => a - b) : k === "tipo" ? vals.sort((a, b) => TIPOS.indexOf(a) - TIPOS.indexOf(b)) : vals.sort((a, b) => a.localeCompare(b));
    const atual = el.value;
    const nome = k === "anos" ? v => `${v} ano${+v > 1 ? "s" : ""}` : v => v;
    el.innerHTML = `<option value="">${r}: todos</option>` + vals.map(v => `<option value="${esc(v)}">${esc(nome(v))} (${cont[v]})</option>`).join("");
    el.value = vals.includes(atual) ? atual : "";
    el.classList.toggle("ativo", !!el.value);
    f[k] = el.value;
  }
  return f;
}
function lerFiltros(pref) { const f = {}; for (const [k] of CAMPOS_FILTRO) { const el = $(`#${pref}-${k}`); if (el) f[k] = el.value; } return f; }
function modalCatalogo(t) {
  const sel = new Map();
  abrirModal(`<h3>Catálogo — adicionar à tabela "${esc(t.titulo || "sem título")}"</h3>
    <div class="filtros"><input class="busca" id="ctBusca" placeholder="Buscar por várias palavras (ex.: premium 3 anos network)...">${barraFiltros("ct")}<button class="btn peq" id="ctLimpar">Limpar filtros</button></div>
    <div class="cat-lista" id="ctLista"></div>
    <div class="rodape"><span class="dica" id="ctInfo" style="margin-right:auto"></span><button class="btn" onclick="fecharModal()">Cancelar</button><button class="btn pri" id="ctAdd" disabled>Adicionar</button></div>`);
  $("#modal").style.maxWidth = "1100px";
  const desenhar = () => {
    const q = $("#ctBusca").value, res = filtrarProdutos(q, atualizarFiltros("ct"));
    let grpAnt = null, html = "";
    for (const p of res.slice(0, 400)) {
      const g = cabecalhoGrupo(p);
      if (g !== grpAnt) { grpAnt = g; html += `<tr class="cat-h"><td colspan="5">${esc(g)}</td></tr>`; }
      const on = sel.has(p.id);
      html += `<tr data-p="${p.id}" class="${on ? "on" : ""}"><td style="width:30px"><input type="checkbox" ${on ? "checked" : ""}></td><td>${realce(p.nome, q)}${p.codigo ? `<div class="dica" style="margin:0">Cód. ${esc(p.codigo)}</div>` : ""}</td><td>${etiquetas(p)}</td><td class="num-r">${fmtBRL(p.valor)}</td><td style="width:80px"><input type="number" min="1" value="${on ? sel.get(p.id) : 1}" style="padding:3px 6px"></td></tr>`;
    }
    $("#ctLista").innerHTML = res.length ? `<table class="lista"><thead><tr><th></th><th>Produto</th><th>Classificação</th><th class="num-r">Valor</th><th>Qtde</th></tr></thead><tbody>${html}</tbody></table>${res.length > 400 ? `<div class="vazio">Mostrando 400 de ${res.length}. Use os filtros.</div>` : ""}` : `<div class="vazio">Nenhum produto com esses filtros.</div>`;
    $("#ctInfo").textContent = `${res.length} produto(s) · ${sel.size} selecionado(s)`;
    $("#ctAdd").disabled = !sel.size; $("#ctAdd").textContent = sel.size ? `Adicionar ${sel.size} item(ns)` : "Adicionar";
  };
  $("#ctBusca").addEventListener("input", desenhar);
  $$("#modal [data-filtro]").forEach(s => s.addEventListener("change", desenhar));
  $("#ctLimpar").onclick = () => { $$("#modal [data-filtro]").forEach(s => s.value = ""); $("#ctBusca").value = ""; desenhar(); };
  $("#ctLista").addEventListener("click", e => {
    const tr = e.target.closest("tr[data-p]"); if (!tr || e.target.type === "number") return;
    const id = tr.dataset.p, qtd = +$("input[type=number]", tr).value || 1;
    if (sel.has(id)) sel.delete(id); else sel.set(id, qtd); desenhar();
  });
  $("#ctLista").addEventListener("input", e => { if (e.target.type !== "number") return; const id = e.target.closest("tr").dataset.p; sel.set(id, +e.target.value || 1); e.target.closest("tr").classList.add("on"); $("input[type=checkbox]", e.target.closest("tr")).checked = true; $("#ctInfo").textContent = `${sel.size} selecionado(s)`; $("#ctAdd").disabled = false; $("#ctAdd").textContent = `Adicionar ${sel.size} item(ns)`; });
  $("#ctAdd").onclick = () => { const n = sel.size; for (const [id, q] of sel) addProduto(t, db.produtos.find(p => p.id === id), q); fecharModal(); toast(n + " item(ns) adicionado(s)."); };
  desenhar(); if (matchMedia("(min-width: 761px)").matches) $("#ctBusca").focus();
}
function acharProduto(txt) { const v = norm(txt).trim(); if (!v) return null; return db.produtos.find(p => norm(p.nome) === v) || null; }

/* ============================ CNPJ (Receita) ============================ */
async function buscarCNPJ(cnpj) {
  const d = soDig(cnpj);
  if (d.length !== 14) throw new Error("Informe um CNPJ com 14 dígitos.");
  let erroMsg = "";
  try {
    const r = await fetch("https://brasilapi.com.br/api/cnpj/v1/" + d);
    if (r.ok) {
      const j = await r.json();
      const tipo = j.descricao_tipo_de_logradouro || "";
      let logr = j.logradouro || "";
      if (tipo && !norm(logr).startsWith(norm(tipo))) logr = tipo + " " + logr;
      const socio = (j.qsa || [])[0];
      return montaCliente({ empresa: j.razao_social, fantasia: j.nome_fantasia, cnpj: d, logradouro: logr, numero: j.numero, complemento: j.complemento, bairro: j.bairro, cep: j.cep, cidade: j.municipio, uf: j.uf, fone: j.ddd_telefone_1, email: j.email, socio: socio && socio.nome_socio });
    }
    erroMsg = r.status === 404 ? "CNPJ não encontrado na Receita Federal." : "BrasilAPI respondeu " + r.status;
    if (r.status === 404) throw new Error(erroMsg);
  } catch (e) { if (/não encontrado/.test(e.message)) throw e; erroMsg = erroMsg || e.message; }
  try {
    const r = await fetch("https://publica.cnpj.ws/cnpj/" + d);
    if (!r.ok) throw new Error(r.status === 429 ? "Muitas consultas seguidas — aguarde 1 minuto." : "Consulta falhou (" + r.status + ")");
    const j = await r.json(); const e = j.estabelecimento || {};
    const socio = (j.socios || [])[0];
    return montaCliente({ empresa: j.razao_social, fantasia: e.nome_fantasia, cnpj: d, logradouro: [e.tipo_logradouro, e.logradouro].filter(Boolean).join(" "), numero: e.numero, complemento: e.complemento, bairro: e.bairro, cep: e.cep, cidade: e.cidade && e.cidade.nome, uf: e.estado && e.estado.sigla, fone: (e.ddd1 || "") + (e.telefone1 || ""), email: e.email, socio: socio && socio.nome });
  } catch (e2) { throw new Error("Não foi possível consultar o CNPJ (" + (e2.message || erroMsg) + "). Verifique a internet ou preencha manualmente."); }
}
function montaCliente(x) {
  const partes = { logradouro: titleCase(x.logradouro), numero: x.numero || "", complemento: x.complemento ? String(x.complemento).toUpperCase().replace(/\s+/g, " ").trim() : "", bairro: titleCase(x.bairro), cep: x.cep || "", cidade: titleCase(x.cidade), uf: (x.uf || "").toUpperCase() };
  return { empresa: titleCase(x.empresa), fantasia: x.fantasia ? titleCase(x.fantasia) : "", cnpj: fmtCNPJ(x.cnpj), contato: x.socio ? titleCase(x.socio) : "", email: (x.email || "").toLowerCase(), telefone: fmtFone(x.fone), endereco: textoEndereco(partes), cidade: partes.cidade, uf: partes.uf, _end: partes };
}

/* ============================ geração do .docx ============================ */
const TAB = "⇥";
function paragrafosDe(texto) {
  return String(texto || "").replace(/\r/g, "").split("\n").map(l => {
    l = l.replace(/\t/g, TAB);
    const m = l.match(/^\*\*(.+?)\*\*(.*)$/);
    return m ? { b: m[1], t: m[2] } : { b: "", t: l };
  });
}
function previewTexto(texto) { return paragrafosDe(texto).map(p => `<p>${p.b ? "<b>" + esc(p.b) + "</b>" : ""}${esc(p.t).replace(new RegExp(TAB, "g"), "&emsp;")}</p>`).join(""); }
function dadosDocumento() {
  const c = db.config, cl = cur.cliente;
  const tabelas = cur.tabelas.map((t, i) => {
    const k = calcTabela(t);
    return {
      titulo: t.titulo,
      itens: t.itens.map(it => ({ qtd: String(it.qtd ?? ""), item: it.nome, unit: fmtBRL(it.valor), total: it.cortesia ? c.textoCortesia : fmtBRL((+it.qtd || 0) * (+it.valor || 0)) })),
      subtotal: fmtBRL(k.sub), labelDesconto: "*Desconto " + fmtPct(t.descPct), desconto: fmtBRL(k.desc), total: fmtBRL(k.total),
      labelParcelas: k.n > 1 ? `*Simulado em ${k.n}x` : "*À vista", parcela: fmtBRL(k.parcela), rotulo: rotuloTabela(t, i),
    };
  });
  let textoTotalGeral = "";
  if (cur.modo === "somatoria") {
    const tg = totalGeral(), n = Math.max(1, +cur.parcelasGeral || 1);
    textoTotalGeral = "Investimento total: " + fmtBRL(tg) + (n > 1 ? `   |   *Simulado em ${n}x de ${fmtBRL(parcelaDe(tg, n))}` : "");
  }
  const descricoes = descLista().filter(d => d.on).map(d => { const x = db.descritivos.find(z => z.id === d.id); return { titulo: x.titulo, paragrafos: paragrafosDe(x.texto) }; });
  return {
    tituloCapa: tituloCapa(), empresa: cl.empresa, contato: cl.contato, numero: cur.numero, validade: dataBR(cur.validade),
    dataExtenso: dataExtenso(cur.data), cidadeEmissao: cur.cidadeEmissao,
    vendedorNome: c.vendedorNome, vendedorFone: c.vendedorFone, vendedorEmail: c.vendedorEmail, filialAceite: c.filialAceite,
    clienteEndereco: cl.endereco, clienteCidadeUF: [cl.cidade, cl.uf].filter(Boolean).join(" - "), clienteTelefone: cl.telefone, clienteCnpj: cl.cnpj,
    formaPagamento: cur.formaPagamento, prazoEntrega: cur.prazoEntrega,
    tabelas, mostrarTotalGeral: cur.modo === "somatoria", textoTotalGeral,
    descricoes, mostrarRequisitos: !!cur.mostrarRequisitos, mostrarMatriz: !!cur.mostrarMatriz,
  };
}
function validar() {
  const f = [];
  if (!cur.cliente.empresa) f.push("nome da empresa");
  if (!cur.cliente.contato) f.push("contato (Para)");
  if (!cur.numero) f.push("número da proposta");
  if (!cur.validade) f.push("validade");
  if (!cur.tabelas.length || !cur.tabelas.some(t => t.itens.length)) f.push("ao menos um item no investimento");
  return f;
}
function montarDocx(b64) {
  const zip = new PizZip(b64ToBuf(b64));
  const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => "" });
  doc.render(dadosDocumento());
  const out = doc.getZip();
  const xml = out.file("word/document.xml").asText();
  out.file("word/document.xml", xml.split(TAB).join('</w:t><w:tab/><w:t xml:space="preserve">'));
  return out.generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", compression: "DEFLATE" });
}
async function gerarDocx() {
  if (!db.modelo) { irAba("config"); return toast("Carregue primeiro o modelo Word (modelo-v2.docx) em Configurações.", true); }
  const falta = validar();
  if (falta.length && !confirm("Faltam: " + falta.join(", ") + ".\n\nGerar mesmo assim?")) return;
  let blob;
  try { blob = montarDocx(db.modelo); } catch (e) {
    console.error(e);
    const det = e.properties && e.properties.errors ? e.properties.errors.map(x => x.properties && x.properties.explanation).join("; ") : e.message;
    return toast("Erro ao montar o documento: " + det, true);
  }
  const nome = nomeArquivo(`Proposta ${cur.numero || "sem numero"} - ${cur.cliente.empresa || "cliente"}.docx`);
  if ("showSaveFilePicker" in window) {
    try {
      const h = await showSaveFilePicker({ suggestedName: nome, id: "propostas", startIn: "documents", types: [{ description: "Documento do Word", accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] } }] });
      const w = await h.createWritable(); await w.write(blob); await w.close(); toast("Proposta gerada: " + h.name);
    } catch (e) { if (e.name === "AbortError") return; baixarBlob(blob, nome); toast("Proposta gerada: " + nome); }
  } else if (navigator.canShare && /Android|iPhone|iPad/i.test(navigator.userAgent)) {
    const f = new File([blob], nome, { type: blob.type });
    if (navigator.canShare({ files: [f] })) { try { await navigator.share({ files: [f], title: nome }); } catch (e) { if (e.name !== "AbortError") baixarBlob(blob, nome); } }
    else baixarBlob(blob, nome);
  } else { baixarBlob(blob, nome); toast("Proposta gerada: " + nome); }
  if (cur.situacao === "Rascunho") cur.situacao = "Enviada";
  await salvarProposta(true);
  renderProposta();
}

/* ============================ propostas salvas ============================ */
async function salvarProposta(silencioso) {
  const total = cur.modo === "somatoria" ? totalGeral() : (cur.tabelas[0] ? calcTabela(cur.tabelas[0]).total : 0);
  const clienteOk = cur.cliente.id && cliPorId(cur.cliente.id);
  const contatoOk = cur.cliente.contatoId && db.contatos.some(x => x.id === cur.cliente.contatoId);
  const row = {
    id: cur.id, numero: vazioNull(cur.numero), cliente_id: clienteOk ? cur.cliente.id : null, contato_id: contatoOk ? cur.cliente.contatoId : null,
    vendedor_id: usuario ? usuario.id : null, origem: cur.origem || null, situacao: cur.situacao || "Rascunho",
    data: cur.data || hojeISO(), validade: cur.validade || null, modo: cur.modo, dados: cur, total,
  };
  const r = await tentar(() => gravar("propostas", row));
  if (!r) return false;
  const i = db.propostas.findIndex(x => x.id === r.id); if (i >= 0) db.propostas[i] = r; else db.propostas.unshift(r);
  salvarRascunho(); renderListaProp();
  if (!silencioso) toast("Proposta salva no banco.");
  return true;
}
function proximaRevisao(n) { const m = String(n || "").match(/^(.*-)(\d+)$/); return m ? m[1] + (+m[2] + 1) : (n ? n + "-2" : ""); }
function corSituacao(s) { return { Ganha: "b-verde", Perdida: "b-vermelho", Cancelada: "b-cinza", "Em negociação": "b-amarelo", Enviada: "b-azul" }[s] || "b-cinza"; }
function renderListaProp() {
  const q = norm($("#buscaProp").value), fs = $("#fpropSit").value, fo = $("#fpropOrig").value;
  const l = db.propostas.filter(p => (!fs || p.situacao === fs) && (!fo || p.origem === fo) && (!q || norm((p.numero || "") + " " + ((p.dados && p.dados.cliente && p.dados.cliente.empresa) || "") + " " + ((cliPorId(p.cliente_id) || {}).nome_fantasia || "")).includes(q)));
  $("#listaProp").innerHTML = l.length ? `<table class="lista"><thead><tr><th>Número</th><th>Cliente</th><th>Tipo</th><th>Situação</th><th>Data</th><th>Validade</th><th class="num-r">Valor</th><th></th></tr></thead><tbody>${l.map(p => {
    const emp = (p.dados && p.dados.cliente && p.dados.cliente.empresa) || "—"; const vencida = p.validade && p.validade < hojeISO() && ["Enviada", "Em negociação"].includes(p.situacao);
    return `<tr data-p="${p.id}"><td><b>${esc(p.numero || "—")}</b></td><td>${esc(emp)}</td><td>${esc(p.origem || "")}</td>
    <td><select data-pa="sit" style="width:auto;padding:3px 6px;font-size:12px">${SITUACOES.map(s => `<option ${s === p.situacao ? "selected" : ""}>${s}</option>`).join("")}</select></td>
    <td>${dataBR(p.data)}</td><td>${dataBR(p.validade)}${vencida ? ' <span class="badge b-vermelho">expirada</span>' : ""}</td><td class="num-r">${fmtBRL(p.total)}</td>
    <td style="white-space:nowrap;text-align:right"><button class="btn peq pri" data-pa="abrir">Abrir</button> <button class="btn peq" data-pa="rev">Nova revisão</button> <button class="btn peq" data-pa="dup">Usar como base</button> <button class="btn peq perigo" data-pa="del">✕</button></td></tr>`; }).join("")}</tbody></table>`
    : `<div class="vazio">Nenhuma proposta. Ao gerar uma proposta ela é salva aqui automaticamente.</div>`;
}
function ligarListaProp() {
  $("#fpropSit").innerHTML = `<option value="">Situação: todas</option>` + SITUACOES.map(s => `<option>${s}</option>`).join("");
  $("#fpropOrig").innerHTML = `<option value="">Tipo: todos</option>` + ORIGENS.map(s => `<option>${s}</option>`).join("");
  for (const id of ["#buscaProp", "#fpropSit", "#fpropOrig"]) $(id).addEventListener(id === "#buscaProp" ? "input" : "change", renderListaProp);
  $("#listaProp").addEventListener("change", async e => {
    if (e.target.dataset.pa !== "sit") return; const id = e.target.closest("[data-p]").dataset.p; const p = db.propostas.find(x => x.id === id);
    const dados = Object.assign({}, p.dados, { situacao: e.target.value });
    const r = await tentar(async () => { const { data, error } = await sb.from("propostas").update({ situacao: e.target.value, dados }).eq("id", id).select().single(); erroSb(error, "atualizar proposta"); return data; });
    if (r) { Object.assign(p, r); if (cur.id === id) { cur.situacao = r.situacao; renderProposta(); } toast("Situação atualizada: " + r.situacao); }
  });
  $("#listaProp").addEventListener("click", async e => {
    const b = e.target.closest("button"); if (!b) return; const id = b.closest("[data-p]").dataset.p; const p = db.propostas.find(x => x.id === id);
    if (b.dataset.pa === "del") {
      if (!confirm("Excluir a proposta " + (p.numero || "") + "?")) return;
      if (!(await tentar(async () => { await apagar("propostas", id); return true; }))) return;
      db.propostas = db.propostas.filter(x => x.id !== id); return renderListaProp();
    }
    const base = Object.assign(novaProposta(), clone(p.dados || {}));
    base.id = p.id; base.situacao = p.situacao; base.origem = p.origem || base.origem;
    if (b.dataset.pa === "abrir") cur = base;
    if (b.dataset.pa === "rev") { cur = base; cur.id = uuid(); cur.numero = proximaRevisao(p.numero); cur.data = hojeISO(); cur.validade = addDias(hojeISO(), db.config.validadeDias); cur.situacao = "Rascunho"; }
    if (b.dataset.pa === "dup") { cur = base; cur.id = uuid(); cur.numero = ""; cur.data = hojeISO(); cur.validade = addDias(hojeISO(), db.config.validadeDias); cur.cliente = clienteVazio(); cur.tituloManual = false; cur.situacao = "Rascunho"; }
    salvarRascunho(); irAba("proposta"); renderProposta();
    toast(b.dataset.pa === "abrir" ? "Proposta aberta." : b.dataset.pa === "rev" ? "Nova revisão criada: " + cur.numero : "Proposta usada como base — escolha o cliente.");
  });
}

/* ============================ renovações (licenças) ============================ */
function renderRenovacoes() {
  const f = $("#frSit").value, q = norm($("#frBusca").value);
  const linhas = db.licencas.map(l => ({ l, s: situacaoLicenca(l), c: cliPorId(l.cliente_id) || {} }))
    .filter(({ l, s, c }) => {
      if (f === "vencida" && s.cod !== "vencida") return false;
      if (f === "alerta" && !(s.cod === "vencida" || s.cod === "avencer")) return false;
      if (/^\d+$/.test(f) && !(s.dias !== null && s.dias >= 0 && s.dias <= +f)) return false;
      return !q || norm([c.nome_fantasia, c.razao_social, c.cnpj_cpf, l.produto_texto, l.numero_serie, (db.produtos.find(p => p.id === l.produto_id) || {}).nome].join(" ")).includes(q);
    })
    .sort((a, b) => (a.l.validade || "9999").localeCompare(b.l.validade || "9999"));
  const nVenc = db.licencas.filter(l => situacaoLicenca(l).cod === "vencida").length, nAv = db.licencas.filter(l => situacaoLicenca(l).cod === "avencer").length;
  $("#frInfo").textContent = `${db.licencas.length} licença(s) cadastrada(s) · ${nVenc} vencida(s) · ${nAv} vencem em até 90 dias`;
  $("#listaRenov").innerHTML = linhas.length ? `<table class="lista"><thead><tr><th>Cliente</th><th>Produto</th><th>Nº de série</th><th>Qtde</th><th>Validade</th><th>Situação</th><th></th></tr></thead><tbody>${linhas.map(({ l, s, c }) => `
    <tr data-l="${l.id}"><td><b>${esc(c.nome_fantasia || c.razao_social || "?")}</b> ${badgeSuporte(c.suporte)}</td>
    <td>${esc(l.produto_texto || (db.produtos.find(p => p.id === l.produto_id) || {}).nome || "")}${l.tipo_licenca ? ` <span class="pill">${esc(l.tipo_licenca)}</span>` : ""}</td>
    <td>${esc(l.numero_serie || "")}</td><td>${esc(l.quantidade)}</td><td>${dataBR(l.validade)}</td><td><span class="badge ${s.cls}">${esc(s.txt)}</span></td>
    <td style="white-space:nowrap;text-align:right"><button class="btn peq pri" data-ra="prop">Proposta</button> <button class="btn peq" data-ra="cli">Cliente</button></td></tr>`).join("")}</tbody></table>`
    : `<div class="vazio">${db.licencas.length ? "Nenhuma licença nessa situação." : "Nenhuma licença cadastrada ainda. Cadastre as licenças (nº de série e validade) dentro de cada cliente, na aba Clientes."}</div>`;
}
function propostaDeLicenca(l) {
  const c = cliPorId(l.cliente_id); if (!c) return;
  const s = situacaoLicenca(l);
  cur = novaProposta(); aplicarCliente(c);
  cur.origem = s.cod === "vencida" ? "Vencida" : "Renovação";
  cur.tituloManual = true; cur.tituloCapa = `Renovação de subscription da empresa ${cur.cliente.empresa}`;
  const base = db.produtos.find(p => p.id === l.produto_id);
  const t = cur.tabelas[0];
  let ren = null;
  if (base && base.familia) ren = db.produtos.find(p => p.ativo && p.familia === base.familia && p.tipo === "Renovação de subscription (ALC)" && (+p.anos || 1) === 1)
    || db.produtos.find(p => p.ativo && p.familia === base.familia && /Renovação/.test(p.tipo));
  if (base && /Renovação/.test(base.tipo)) ren = ren || base;
  if (ren) addProduto(t, ren, l.quantidade || 1);
  else t.itens.push({ qtd: l.quantidade || 1, prodId: null, nome: "Renovação - " + (l.produto_texto || (base && base.nome) || "licença") + (l.numero_serie ? " (S/N " + l.numero_serie + ")" : ""), valor: 0, cortesia: false });
  if (s.cod === "vencida") t.itens.push({ qtd: s.meses, prodId: null, nome: `Taxa de retroatividade - ${s.meses} ${s.meses > 1 ? "meses" : "mês"} vencido(s) (desde ${dataBR(l.validade)})`, valor: 0, cortesia: false });
  salvarRascunho(); irAba("proposta"); renderProposta();
  toast(s.cod === "vencida" ? `Licença vencida há ${s.meses} mês(es): confira o produto e preencha o valor da retroatividade.` : "Proposta de renovação criada — confira o produto e o valor.");
}
function ligarRenovacoes() {
  $("#frSit").addEventListener("change", renderRenovacoes); $("#frBusca").addEventListener("input", renderRenovacoes);
  $("#listaRenov").addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return; const l = db.licencas.find(x => x.id === b.closest("[data-l]").dataset.l);
    if (b.dataset.ra === "prop") propostaDeLicenca(l);
    if (b.dataset.ra === "cli") { const c = cliPorId(l.cliente_id); if (c) modalCliente(c, "licencas"); }
  });
}

/* ============================ clientes ============================ */
function renderListaCli() {
  const q = norm($("#buscaCli").value), fs = $("#fcStatus").value, fsu = $("#fcSuporte").value, fn = $("#fcNivel").value;
  const l = db.clientes.filter(c => (!fs || c.status_licenciamento === fs) && (!fsu || c.suporte === fsu) && (!fn || (fn === "(vazio)" ? !c.nivel : c.nivel === fn))
    && (!q || norm([c.nome_fantasia, c.razao_social, c.cnpj_cpf, soDig(c.cnpj_cpf), c.cidade, c.uf, c.codigo_crm, contatosDe(c.id).map(x => x.nome + " " + (x.email || "")).join(" ")].join(" ")).includes(q)))
    .sort((a, b) => (a.nome_fantasia || "").localeCompare(b.nome_fantasia || ""));
  const max = 300;
  $("#contCli").textContent = `${l.length} de ${db.clientes.length} cliente(s)` + (l.length > max ? ` — mostrando os ${max} primeiros, refine a busca` : "");
  $("#listaCli").innerHTML = l.length ? `<table class="lista"><thead><tr><th>Cliente</th><th>CNPJ</th><th>Cidade</th><th>Situação</th><th></th></tr></thead><tbody>${l.slice(0, max).map(c => `
    <tr data-c-id="${c.id}"><td><b>${esc(c.nome_fantasia)}</b>${c.razao_social ? `<div class="dica" style="margin:0">${esc(c.razao_social)}${c.codigo_crm ? " · CRM " + c.codigo_crm : ""}</div>` : ""}</td>
    <td style="white-space:nowrap">${esc(c.cnpj_cpf || "")}</td><td>${esc([c.cidade, c.uf].filter(Boolean).join(" - "))}</td>
    <td>${badgeSuporte(c.suporte)} ${badgeNivel(c.nivel)} ${c.status_licenciamento === "Inativo" ? '<span class="badge b-cinza">Inativo</span>' : ""}</td>
    <td style="white-space:nowrap;text-align:right"><button class="btn peq pri" data-ca="usar">Nova proposta</button> <button class="btn peq" data-ca="edit">Abrir</button></td></tr>`).join("")}</tbody></table>` : `<div class="vazio">Nenhum cliente.</div>`;
}
function preencherFiltrosCli() {
  const sup = [...new Set(db.clientes.map(c => c.suporte).filter(Boolean))].sort();
  const niv = [...new Set(db.clientes.map(c => c.nivel).filter(Boolean))].sort();
  const a = $("#fcSuporte").value, b = $("#fcNivel").value;
  $("#fcSuporte").innerHTML = `<option value="">Suporte: todos</option>` + sup.map(s => `<option>${esc(s)}</option>`).join("");
  $("#fcNivel").innerHTML = `<option value="">Nível: todos</option>` + niv.map(s => `<option>${esc(s)}</option>`).join("") + `<option value="(vazio)">Sem nível</option>`;
  $("#fcSuporte").value = a; $("#fcNivel").value = b;
}
const CAMPOS_CLIENTE = [
  ["nome_fantasia", "Nome fantasia", "texto", "span2"], ["codigo_crm", "Código CRM", "numero"], ["ativo", "Cadastro ativo", "bool"],
  ["razao_social", "Razão social", "texto", "span2"], ["cnpj_cpf", "CNPJ / CPF", "texto"], ["inscricao_estadual", "Inscrição estadual", "texto"],
  ["cidade", "Cidade", "texto"], ["uf", "UF", "texto"], ["site", "Site", "texto", "span2"],
  ["classificacao", "Classificação", ["", "SMB", "Mid-Market", "Enterprise"]], ["nivel", "Nível", ["", "Ouro Plus", "Ouro", "Prata"]],
  ["status_licenciamento", "Status licenciamento", ["", "Ativo", "Inativo"]], ["suporte", "Suporte", ["", "Verde (liberado)", "Amarelo (restrição)", "Vermelho (bloqueado)", "Cinza (sem suporte)"]],
  ["em_risco_para", "Em risco para", "texto"], ["bloquear_suporte_para", "Bloquear suporte para", "texto"],
  ["pontos_atencao", "Pontos de atenção", "area", "span4"], ["observacoes", "Observações", "area", "span4"],
];
function campoHtml(pref, [k, rot, tipo, cls], v) {
  const id = pref + k, st = cls === "span4" ? ' style="grid-column:1/-1"' : "";
  const wrap = h => `<div class="${cls === "span4" ? "" : cls || ""}"${st}>${h}</div>`;
  if (tipo === "bool") return wrap(`<label>&nbsp;</label><label class="chk"><input type="checkbox" id="${id}" ${v ? "checked" : ""}> ${esc(rot)}</label>`);
  if (Array.isArray(tipo)) { const ops = [...tipo]; if (v && !ops.includes(v)) ops.push(v); return wrap(`<label>${esc(rot)}</label><select id="${id}">${ops.map(o => `<option value="${esc(o)}" ${o === (v || "") ? "selected" : ""}>${esc(o || "—")}</option>`).join("")}</select>`); }
  if (tipo === "area") return wrap(`<label>${esc(rot)}</label><textarea id="${id}" rows="2" style="min-height:50px">${esc(v || "")}</textarea>`);
  if (tipo === "lista") return wrap(`<label>${esc(rot)}</label><input id="${id}" list="${cls}" value="${esc(v || "")}">`);
  const t = tipo === "numero" ? "number" : tipo === "data" ? "date" : "text";
  return wrap(`<label>${esc(rot)}</label><input type="${t}" id="${id}" value="${esc(v ?? "")}">`);
}
function lerCampo(pref, [k, , tipo]) {
  const el = $("#" + pref + k); if (!el) return undefined;
  if (tipo === "bool") return el.checked;
  if (tipo === "numero") return el.value === "" ? null : +el.value;
  return vazioNull(el.value);
}
/* ---- sub-cadastros: contatos, endereços, telefones, licenças ---- */
const SUB = {
  contatos: { titulo: "Contatos", campos: [["nome", "Nome", "texto", "span2"], ["cargo", "Cargo", "texto"], ["email", "E-mail", "texto", "span2"], ["telefone", "Telefone", "texto"], ["celular", "Celular", "texto"], ["ramal", "Ramal", "texto"],
      ["classificacao", "Classificação", ["Não classificado", "Decisor", "Financeiro / Compras", "Usuário SOLIDWORKS", "Usuário 3DEXPERIENCE", "Usuário QC", "TI", "Outro"]], ["decisor", "Decisor", "bool"], ["ativo", "Ativo", "bool"], ["observacoes", "Observações", "area", "span4"]],
    novo: { ativo: true, decisor: false, classificacao: "Não classificado" }, obrig: "nome",
    resumo: x => `<b>${esc(x.nome)}</b>${x.decisor ? ' <span class="badge b-amarelo">decisor</span>' : ""}${x.ativo === false ? ' <span class="badge b-cinza">inativo</span>' : ""}<div class="dica" style="margin:0">${esc([x.cargo, x.email, x.telefone, x.celular].filter(Boolean).join(" · "))}</div>` },
  enderecos: { titulo: "Endereços", campos: [["logradouro", "Logradouro (rua / endereço completo)", "texto", "span2"], ["numero", "Número", "texto"], ["complemento", "Complemento", "texto"], ["bairro", "Bairro", "texto"], ["cep", "CEP", "texto"], ["cidade", "Cidade", "texto"], ["uf", "UF", "texto"],
      ["tipo", "Tipo", ["Comercial", "Cobrança", "Entrega", "Outro"]], ["principal", "Principal (vai na proposta)", "bool"]],
    novo: { tipo: "Comercial", principal: false }, obrig: "logradouro",
    resumo: x => `${esc(textoEndereco(x))}${x.principal ? ' <span class="badge b-azul">principal</span>' : ""}<div class="dica" style="margin:0">${esc([x.cep, [x.cidade, x.uf].filter(Boolean).join(" - "), x.tipo].filter(Boolean).join(" · "))}</div>` },
  telefones: { titulo: "Telefones", campos: [["tipo", "Tipo", ["Comercial", "Celular", "Financeiro", "Outro"]], ["ddd", "DDD", "texto"], ["numero", "Número", "texto"], ["ramal", "Ramal", "texto"], ["observacoes", "Observações", "texto", "span2"]],
    novo: { tipo: "Comercial", ddi: "55" }, obrig: "numero",
    resumo: x => `<b>${esc(textoTelefone(x))}</b> <span class="dica">${esc([x.tipo, x.observacoes].filter(Boolean).join(" · "))}</span>` },
  licencas: { titulo: "Licenças", campos: [["produto_texto", "Produto (como está no CRM)", "texto", "span2"], ["numero_serie", "Nº de série", "texto"], ["_produto", "Produto do catálogo (opcional — ajuda a sugerir a renovação)", "lista", "dlProdutos"],
      ["quantidade", "Quantidade", "numero"], ["validade", "Validade (fim da subscription)", "data"], ["tipo_licenca", "Tipo", ["", "Local", "Em rede"]], ["fornecedor", "Fornecedor", "texto"],
      ["status", "Status", "texto"], ["negocio", "Negócio", "texto"], ["temporario", "Temporária", "bool"], ["observacoes", "Observações", "area", "span4"]],
    novo: { quantidade: 1, fornecedor: "Dassault Systèmes" }, obrig: "produto_texto",
    resumo: x => { const s = situacaoLicenca(x); return `<b>${esc(x.produto_texto || (db.produtos.find(p => p.id === x.produto_id) || {}).nome || "")}</b> <span class="badge ${s.cls}">${esc(s.txt)}</span><div class="dica" style="margin:0">${esc([x.numero_serie && "S/N " + x.numero_serie, "Qtde " + (x.quantidade || 1), x.validade && "validade " + dataBR(x.validade), x.tipo_licenca].filter(Boolean).join(" · "))}</div>`; } },
};
function renderSub(tabela, clienteId, editId) {
  const cfg = SUB[tabela], box = $(`#sub-${tabela}`); if (!box) return;
  const itens = db[tabela].filter(x => x.cliente_id === clienteId);
  $(".cab .n", box.parentElement).textContent = `${cfg.titulo} (${itens.length})`;
  const form = (x) => {
    const v = Object.assign({}, x); if (tabela === "licencas") v._produto = (db.produtos.find(p => p.id === x.produto_id) || {}).nome || "";
    return `<div class="sub-form" data-form="${x.id || ""}"><div class="grid g4">${cfg.campos.map(c => campoHtml("sf-", c, v[c[0]])).join("")}</div>
      <div class="rodape" style="margin-top:8px"><button class="btn peq" data-sa="cancelar">Cancelar</button><button class="btn peq pri" data-sa="salvar">Salvar</button></div></div>`;
  };
  box.innerHTML = itens.map(x => editId === x.id ? form(x) : `<div class="sub-item" data-id="${x.id}"><div class="t">${cfg.resumo(x)}</div><button class="btn peq" data-sa="editar">Editar</button><button class="btn peq perigo" data-sa="excluir">✕</button></div>`).join("")
    + (editId === "novo" ? form(Object.assign({}, cfg.novo)) : "")
    + (!itens.length && editId !== "novo" ? `<div class="dica">Nenhum registro.</div>` : "");
}
function ligarSub(tabela, clienteId) {
  const box = $(`#sub-${tabela}`), cfg = SUB[tabela];
  $(`[data-subnovo="${tabela}"]`).onclick = () => renderSub(tabela, clienteId, "novo");
  box.onclick = async e => {
    const b = e.target.closest("button[data-sa]"); if (!b) return;
    const a = b.dataset.sa, item = b.closest("[data-id]"), fEl = b.closest("[data-form]");
    if (a === "editar") return renderSub(tabela, clienteId, item.dataset.id);
    if (a === "cancelar") return renderSub(tabela, clienteId);
    if (a === "excluir") {
      if (!confirm("Excluir este registro?")) return;
      if ((await tentar(async () => { await apagar(tabela, item.dataset.id); return true; }))) { db[tabela] = db[tabela].filter(x => x.id !== item.dataset.id); renderSub(tabela, clienteId); posAlteracaoCliente(); }
      return;
    }
    if (a === "salvar") {
      const id = fEl.dataset.form, antigo = id ? db[tabela].find(x => x.id === id) : null;
      const row = Object.assign({}, antigo || { id: uuid(), cliente_id: clienteId });
      for (const c of cfg.campos) if (c[0] !== "_produto") row[c[0]] = lerCampo("sf-", c);
      if (!row[cfg.obrig]) return toast("Preencha o campo obrigatório.", true);
      if (tabela === "licencas") { const nm = $("#sf-_produto").value.trim(); const p = nm ? acharProduto(nm) : null; if (nm && !p) return toast("Produto do catálogo não encontrado — escolha da lista ou deixe vazio.", true); row.produto_id = p ? p.id : null; if (!row.quantidade) row.quantidade = 1; }
      if (tabela === "enderecos" && row.principal) for (const e2 of enderecosDe(clienteId)) if (e2.id !== row.id && e2.principal) { await tentar(() => gravar("enderecos", { ...e2, principal: false })); e2.principal = false; }
      delete row.criado_em;
      const r = await tentar(() => gravar(tabela, row));
      if (r) { substituir(db[tabela], r); renderSub(tabela, clienteId); posAlteracaoCliente(); toast("Salvo."); }
    }
  };
}
function posAlteracaoCliente() { renderListaCli(); renderRenovacoes(); renderDatalists(); if (cur.cliente.id) renderAlertaCliente(cliPorId(cur.cliente.id)); }
function modalCliente(c, focar) {
  const novo = !c; c = c ? { ...c } : { id: uuid(), ativo: true, status_licenciamento: "Ativo" };
  abrirModal(`<h3>${novo ? "Novo cliente" : esc(c.nome_fantasia || c.razao_social)}</h3>
    <div class="linha" style="margin-bottom:12px"><div class="cresce"><label>Buscar dados na Receita pelo CNPJ</label><input id="mCnpjBusca" value="${esc(c.cnpj_cpf || "")}" inputmode="numeric"></div><button class="btn pri" id="mBusca">Buscar CNPJ</button></div>
    <div class="grid g4">${CAMPOS_CLIENTE.map(f => campoHtml("mc-", f, c[f[0]])).join("")}</div>
    <div class="rodape" style="justify-content:space-between">
      <span>${novo ? "" : `<button class="btn perigo" id="mExcluir">Excluir cliente</button>`}</span>
      <span class="linha"><button class="btn" onclick="fecharModal()">Fechar</button>${novo ? "" : `<button class="btn" id="mProp">Nova proposta</button>`}<button class="btn pri" id="mSalvar">Salvar dados</button></span>
    </div>
    ${novo ? `<div class="dica">Salve os dados para poder incluir contatos, endereços, telefones e licenças.</div>` : Object.keys(SUB).map(t => `
    <div class="sub-sec" id="sec-${t}"><div class="cab"><span class="cresce n">${SUB[t].titulo}</span><button class="btn peq pri" data-subnovo="${t}">+ Adicionar</button></div><div class="cor" id="sub-${t}"></div></div>`).join("")}`);
  $("#modal").style.maxWidth = "980px";
  if (!novo) for (const t of Object.keys(SUB)) { renderSub(t, c.id); ligarSub(t, c.id); }
  if (focar) setTimeout(() => { const s = $("#sec-" + focar); if (s) s.scrollIntoView({ behavior: "smooth" }); }, 50);
  let endReceita = null, contatoReceita = null;
  $("#mBusca").onclick = async () => {
    const b = $("#mBusca"); b.disabled = true; b.innerHTML = '<span class="spin"></span> Buscando';
    try {
      const r = await buscarCNPJ($("#mCnpjBusca").value);
      $("#mc-razao_social").value = r.empresa; $("#mc-cnpj_cpf").value = r.cnpj; $("#mc-cidade").value = r.cidade; $("#mc-uf").value = r.uf;
      if (!$("#mc-nome_fantasia").value) $("#mc-nome_fantasia").value = r.fantasia || r.empresa;
      endReceita = r._end; contatoReceita = r.contato ? { nome: r.contato, email: r.email, telefone: r.telefone } : null;
      toast("Dados da Receita carregados" + (novo ? " — o endereço será incluído ao salvar." : " — ao salvar, o endereço da Receita vira o principal."));
    } catch (e) { toast(e.message, true); } finally { b.disabled = false; b.textContent = "Buscar CNPJ"; }
  };
  $("#mSalvar").onclick = () => tentar(async () => {
    const row = { ...c }; delete row.criado_em; delete row.atualizado_em;
    for (const f of CAMPOS_CLIENTE) row[f[0]] = lerCampo("mc-", f);
    if (!row.nome_fantasia) row.nome_fantasia = row.razao_social;
    if (!row.nome_fantasia) return toast("Informe o nome fantasia ou a razão social.", true);
    if (row.cnpj_cpf) { row.cnpj_cpf = fmtCNPJ(row.cnpj_cpf); const d = soDig(row.cnpj_cpf); row.tipo_documento = d.length === 14 ? "PJ (CNPJ)" : d.length === 11 ? "PF (CPF)" : row.tipo_documento; if (d.length === 14) row.cnpj_raiz = row.cnpj_cpf.slice(0, 10); }
    if (row.uf) row.uf = row.uf.toUpperCase();
    const dup = row.cnpj_cpf && db.clientes.find(x => x.id !== row.id && soDig(x.cnpj_cpf) === soDig(row.cnpj_cpf));
    if (dup && !confirm(`Já existe um cliente com esse CNPJ: ${dup.nome_fantasia}. Salvar mesmo assim?`)) return;
    const r = await gravar("clientes", row); substituir(db.clientes, r);
    if (endReceita) {
      for (const e2 of enderecosDe(r.id)) if (e2.principal) { await gravar("enderecos", { ...e2, principal: false }); e2.principal = false; }
      substituir(db.enderecos, await gravar("enderecos", { id: uuid(), cliente_id: r.id, tipo: "Comercial", principal: true, ...Object.fromEntries(Object.entries(endReceita).map(([k, v]) => [k, vazioNull(v)])) }));
    }
    if (novo && contatoReceita) substituir(db.contatos, await gravar("contatos", { id: uuid(), cliente_id: r.id, nome: contatoReceita.nome, email: vazioNull(contatoReceita.email), telefone: vazioNull(contatoReceita.telefone), observacoes: "Sócio (dados da Receita)" }));
    preencherFiltrosCli(); posAlteracaoCliente(); toast("Cliente salvo.");
    if (novo) modalCliente(r); else { c = r; endReceita = null; }
  });
  if (!novo) {
    $("#mProp").onclick = () => { cur = novaProposta(); aplicarCliente(cliPorId(c.id)); salvarRascunho(); fecharModal(); irAba("proposta"); renderProposta(); toast("Nova proposta para " + (c.nome_fantasia || c.razao_social)); };
    $("#mExcluir").onclick = () => tentar(async () => {
      const nProp = db.propostas.filter(p => p.cliente_id === c.id).length;
      if (nProp) return toast(`Este cliente tem ${nProp} proposta(s) — não pode ser excluído. Desmarque "Cadastro ativo" para inativá-lo.`, true);
      if (!confirm(`Excluir DEFINITIVAMENTE ${c.nome_fantasia} e todos os seus contatos, endereços, telefones e licenças?`)) return;
      await apagar("clientes", c.id);
      db.clientes = db.clientes.filter(x => x.id !== c.id); for (const t of Object.keys(SUB)) db[t] = db[t].filter(x => x.cliente_id !== c.id);
      fecharModal(); posAlteracaoCliente(); toast("Cliente excluído.");
    });
  }
}
function ligarClientes() {
  for (const id of ["#buscaCli", "#fcStatus", "#fcSuporte", "#fcNivel"]) $(id).addEventListener(id === "#buscaCli" ? "input" : "change", renderListaCli);
  $("#btnNovoCli").addEventListener("click", () => modalCliente(null));
  $("#listaCli").addEventListener("click", e => {
    const b = e.target.closest("button"); const tr = e.target.closest("[data-c-id]"); if (!tr) return;
    const c = cliPorId(tr.dataset.cId);
    if (b && b.dataset.ca === "usar") { cur = novaProposta(); aplicarCliente(c); salvarRascunho(); irAba("proposta"); renderProposta(); toast("Nova proposta para " + (c.nome_fantasia || c.razao_social)); }
    else modalCliente(c);
  });
}

/* ============================ produtos ============================ */
function renderListaProd() {
  const q = $("#buscaProd").value;
  const l = filtrarProdutos(q, atualizarFiltros("fp"), true);
  const max = 300;
  $("#contProd").textContent = `${l.length} de ${db.produtos.length} produto(s)` + (l.length > max ? ` — mostrando ${max}, use a busca ou os filtros` : "");
  let grpAnt = null, linhas = "";
  for (const p of l.slice(0, max)) {
    const g = cabecalhoGrupo(p);
    if (g !== grpAnt) { grpAnt = g; linhas += `<tr class="cat-h"><td colspan="5">${esc(g)}</td></tr>`; }
    linhas += `<tr data-pr="${p.id}"><td><b>${realce(p.nome, q)}</b>${p.codigo ? `<div class="dica" style="margin:0">Cód. ${esc(p.codigo)}</div>` : ""}</td><td>${etiquetas(p) || "—"}</td><td class="num-r">${fmtBRL(p.valor)}</td>
    <td>${(p.descritivos || []).map(id => { const d = db.descritivos.find(x => x.id === id); return d ? `<span class="pill" title="${esc(d.titulo)}">${esc(d.titulo.length > 26 ? d.titulo.slice(0, 26) + "…" : d.titulo)}</span>` : ""; }).join(" ") || "—"}</td>
    <td style="white-space:nowrap;text-align:right"><button class="btn peq" data-pa="edit">Editar</button> <button class="btn peq" data-pa="dup">Duplicar</button></td></tr>`;
  }
  $("#listaProd").innerHTML = l.length ? `<table class="lista"><thead><tr><th>Produto (texto do item na tabela)</th><th>Classificação</th><th class="num-r">Valor unitário</th><th>Descritivos</th><th></th></tr></thead><tbody>${linhas}</tbody></table>` : `<div class="vazio">Nenhum produto com esses filtros.</div>`;
  $("#gruposTxt").value = db.grupos.join("\n");
}
function opcoesSelect(lista, atual, vazio) {
  const l = [...lista]; if (atual && !l.includes(atual)) l.push(atual);
  return `<option value="">${vazio}</option>` + l.map(v => `<option value="${esc(v)}" ${v === atual ? "selected" : ""}>${esc(v)}</option>`).join("");
}
async function salvarProdutoDb(p) {
  const r = await gravar("produtos", prodParaRow(p));
  const novo = prodDeRow(r); novo.descritivos = p.descritivos || [];
  const { error } = await sb.from("produto_descritivos").delete().eq("produto_id", r.id); erroSb(error, "atualizar descritivos do produto");
  if (novo.descritivos.length) { const { error: e2 } = await sb.from("produto_descritivos").insert(novo.descritivos.map((d, i) => ({ produto_id: r.id, descritivo_id: d, ordem: i }))); erroSb(e2, "vincular descritivos"); }
  substituir(db.produtos, novo); return novo;
}
function modalProduto(p, ehNovo) {
  const novo = ehNovo || !p;
  if (!p) { const f = lerFiltros("fp"); p = { id: uuid(), nome: "", valor: 0, grupo: "", codigo: "", negocio: f.negocio || "", linha: f.linha || "", subgrupo: f.subgrupo || "", tipo: f.tipo || "", anos: f.anos || "", meses: "", licenca: f.licenca || "", familia: "", base: "", conferir: "", ativo: true, descritivos: [], cortesia: false }; }
  else p = clone(p);
  abrirModal(`<h3>${novo ? "Novo produto" : "Editar produto"}</h3>
    <div class="grid g3">
      <div><label>Código CRM</label><input id="mpCod" value="${esc(p.codigo)}" inputmode="numeric"></div>
      <div class="span2"><label>Descrição (exatamente como aparece no item da tabela)</label><input id="mpNome" value="${esc(p.nome)}"></div>
      <div><label>Negócio</label><input id="mpNeg" list="dlNegocios" value="${esc(p.negocio)}" placeholder="SOLIDWORKS"></div>
      <div><label>Grupo</label><input id="mpLinha" list="dlLinhas" value="${esc(p.linha)}" placeholder="SOLIDWORKS Simulation"></div>
      <div><label>Subgrupo</label><input id="mpSub" list="dlSubgrupos" value="${esc(p.subgrupo)}"></div>
      <div><label>Tipo</label><select id="mpTipo">${opcoesSelect(tiposExistentes(), p.tipo, "—")}</select></div>
      <div><label>Anos de subscription</label><select id="mpAnos">${opcoesSelect(["1", "2", "3", "4", "5"], String(p.anos || ""), "—")}</select></div>
      <div><label>Meses</label><select id="mpMeses">${opcoesSelect(["1", "3"], String(p.meses || ""), "—")}</select></div>
      <div><label>Licença</label><select id="mpLic">${opcoesSelect(LICENCAS, p.licenca, "—")}</select></div>
      <div><label>Valor unitário (R$)</label><input type="number" step="0.01" id="mpValor" value="${esc(p.valor)}"></div>
      <div><label>Família (liga licença ↔ renovação)</label><input id="mpFam" value="${esc(p.familia)}"></div>
      <div class="span2"><label>Título da tabela na proposta (vazio = automático pelo tipo)</label><input id="mpGrupo" list="dlGrupos" value="${esc(p.grupo)}" placeholder="${esc(tituloPadraoTabela(p))}"></div>
      <div><label>Conferir (observação)</label><input id="mpConf" value="${esc(p.conferir)}"></div>
      <div class="span3 linha" style="gap:22px"><label class="chk"><input type="checkbox" id="mpCort" ${p.cortesia ? "checked" : ""}> Normalmente oferecido como cortesia</label>
        <label class="chk"><input type="checkbox" id="mpAtivo" ${p.ativo ? "checked" : ""}> Ativo (aparece na busca da proposta)</label></div>
      <div class="span3"><label>Descritivos carregados na "Descrição da solução" quando este produto estiver na proposta</label>
        <div class="multi">${db.descritivos.map(d => `<label><input type="checkbox" value="${d.id}" ${p.descritivos.includes(d.id) ? "checked" : ""}> ${esc(d.titulo)}</label>`).join("") || "<i>Nenhum descritivo cadastrado.</i>"}</div></div>
    </div>
    <div class="rodape" style="flex-wrap:wrap"><button class="btn" id="mpAuto" style="margin-right:auto" title="Preenche tipo, anos e licença a partir da descrição">Classificar pela descrição</button>${novo ? "" : '<button class="btn perigo" id="mpDel">Excluir</button>'}<button class="btn" onclick="fecharModal()">Cancelar</button><button class="btn pri" id="mpSalvar">Salvar</button></div>`);
  $("#modal").style.maxWidth = "900px";
  const ler = () => {
    p.codigo = $("#mpCod").value.trim(); p.nome = $("#mpNome").value.trim(); p.negocio = $("#mpNeg").value.trim(); p.linha = $("#mpLinha").value.trim(); p.subgrupo = $("#mpSub").value.trim();
    p.tipo = $("#mpTipo").value; p.anos = $("#mpAnos").value; p.meses = $("#mpMeses").value; p.licenca = $("#mpLic").value; p.valor = +$("#mpValor").value || 0;
    p.familia = $("#mpFam").value.trim(); p.grupo = $("#mpGrupo").value.trim(); p.conferir = $("#mpConf").value.trim();
    p.cortesia = $("#mpCort").checked; p.ativo = $("#mpAtivo").checked; p.descritivos = $$(".multi input:checked", $("#modal")).map(i => i.value);
  };
  $("#mpAuto").onclick = () => {
    const c = classificarPeloNome($("#mpNome").value);
    if (c.tipo) $("#mpTipo").innerHTML = opcoesSelect(tiposExistentes(), c.tipo, "—");
    if (c.anos) $("#mpAnos").value = c.anos; if (c.licenca) $("#mpLic").value = c.licenca;
    toast("Classificação sugerida pela descrição — confira.");
  };
  $("#mpSalvar").onclick = () => tentar(async () => {
    ler(); if (!p.nome) return toast("Informe a descrição.", true);
    await salvarProdutoDb(p);
    if (p.grupo && !db.grupos.includes(p.grupo)) { db.grupos.push(p.grupo); await gravarConfig("grupos_tabela", { lista: db.grupos }); }
    fecharModal(); renderListaProd(); renderDatalists(); renderListaDesc(); toast("Produto salvo.");
  });
  if (!novo) $("#mpDel").onclick = () => tentar(async () => {
    if (!confirm("Excluir o produto " + p.nome + "?\n\nDica: se ele só não é mais vendido, prefira desmarcar \"Ativo\".")) return;
    await apagar("produtos", p.id); db.produtos = db.produtos.filter(x => x.id !== p.id); fecharModal(); renderListaProd(); renderDatalists(); toast("Produto excluído.");
  });
}
function ligarProdutos() {
  $("#fpFiltros").innerHTML = barraFiltros("fp") + `<button class="btn peq" id="fpLimpar">Limpar filtros</button>`;
  $("#buscaProd").addEventListener("input", renderListaProd);
  $$("#fpFiltros [data-filtro]").forEach(s => s.addEventListener("change", renderListaProd));
  $("#fpLimpar").addEventListener("click", () => { $$("#fpFiltros [data-filtro]").forEach(s => s.value = ""); $("#buscaProd").value = ""; renderListaProd(); });
  $("#btnNovoProd").addEventListener("click", () => modalProduto(null));
  $("#btnSalvarGrupos").addEventListener("click", () => tentar(async () => { db.grupos = $("#gruposTxt").value.split("\n").map(s => s.trim()).filter(Boolean); await gravarConfig("grupos_tabela", { lista: db.grupos }); renderDatalists(); toast("Títulos salvos."); }));
  $("#listaProd").addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return; const id = b.closest("[data-pr]").dataset.pr; const p = db.produtos.find(x => x.id === id);
    if (b.dataset.pa === "edit") modalProduto(p);
    if (b.dataset.pa === "dup") modalProduto(Object.assign(clone(p), { id: uuid(), nome: p.nome + " (cópia)", codigo: "" }), true);
  });
}

/* ============================ descritivos ============================ */
function renderListaDesc() {
  const q = norm($("#buscaDesc").value);
  const l = db.descritivos.filter(d => !q || norm(d.titulo + " " + d.texto).includes(q));
  $("#listaDesc").innerHTML = l.length ? `<table class="lista"><thead><tr><th>Título</th><th>Usado nos produtos</th><th></th></tr></thead><tbody>${l.map(d => {
    const usados = db.produtos.filter(p => (p.descritivos || []).includes(d.id)).length;
    return `<tr data-de="${d.id}"><td><b>${esc(d.titulo)}</b><div class="dica" style="margin:2px 0 0">${esc(d.texto.replace(/\s+/g, " ").slice(0, 140))}…</div></td><td>${usados || "—"}</td>
    <td style="white-space:nowrap;text-align:right"><button class="btn peq" data-da="edit">Editar</button> <button class="btn peq" data-da="dup">Duplicar</button> <button class="btn peq perigo" data-da="del">✕</button></td></tr>`; }).join("")}</tbody></table>` : `<div class="vazio">Nenhum descritivo.</div>`;
}
function modalDesc(d, ehNovo) {
  const novo = ehNovo || !d; d = d ? clone(d) : { id: uuid(), titulo: "", texto: "" };
  const prodsVinc = db.produtos.filter(p => p.descritivos.includes(d.id));
  abrirModal(`<h3>${novo ? "Novo descritivo" : "Editar descritivo"}</h3>
    <div class="grid">
      <div><label>Título (aparece em negrito na proposta)</label><input id="mdTit" value="${esc(d.titulo)}"></div>
      <div><label>Texto</label><textarea id="mdTxt" rows="14">${esc(d.texto)}</textarea>
      <div class="dica">Cada linha vira um parágrafo; linha em branco = espaço entre parágrafos. Para negrito no início da linha use <code>**Objetivo:** texto</code>. Para marcadores, comece a linha com <code>•</code> seguido de TAB (ou cole do Word).</div></div>
      <div><label>Produtos que carregam este descritivo automaticamente (${prodsVinc.length})</label>
        <div class="multi" id="mdProds">${prodsVinc.map(p => `<label><input type="checkbox" value="${p.id}" checked> ${esc(p.nome)}</label>`).join("") || "<i>Nenhum.</i>"}</div>
        <div class="linha" style="margin-top:6px"><input id="mdAddProd" list="dlProdutos" class="cresce" placeholder="Vincular a outro produto: digite o nome e escolha..."><button class="btn peq" id="mdAddBtn">Vincular</button></div></div>
      <div><label>Pré-visualização</label><div class="preview-desc" id="mdPrev"></div></div>
    </div>
    <div class="rodape"><button class="btn" onclick="fecharModal()">Cancelar</button><button class="btn pri" id="mdSalvar">Salvar</button></div>`);
  const prev = () => $("#mdPrev").innerHTML = `<p><b>${esc($("#mdTit").value)}</b></p><p></p>` + previewTexto($("#mdTxt").value);
  $("#mdTxt").addEventListener("input", prev); $("#mdTit").addEventListener("input", prev); prev();
  $("#mdTxt").addEventListener("keydown", e => { if (e.key === "Tab") { e.preventDefault(); const t = e.target, s = t.selectionStart; t.value = t.value.slice(0, s) + "\t" + t.value.slice(t.selectionEnd); t.selectionStart = t.selectionEnd = s + 1; prev(); } });
  $("#mdAddBtn").onclick = () => {
    const p = acharProduto($("#mdAddProd").value); if (!p) return toast("Escolha um produto da lista.", true);
    const box = $("#mdProds"); if ($(`input[value="${p.id}"]`, box)) return; if (!$("input", box)) box.innerHTML = "";
    box.insertAdjacentHTML("beforeend", `<label><input type="checkbox" value="${p.id}" checked> ${esc(p.nome)}</label>`); $("#mdAddProd").value = "";
  };
  $("#mdSalvar").onclick = () => tentar(async () => {
    d.titulo = $("#mdTit").value.trim(); d.texto = $("#mdTxt").value.replace(/\s+$/, "");
    if (!d.titulo) return toast("Informe o título.", true);
    const r = await gravar("descritivos", { id: d.id, titulo: d.titulo, texto: d.texto });
    substituir(db.descritivos, { id: r.id, titulo: r.titulo, texto: r.texto || "" });
    const marcados = new Set($$("#mdProds input:checked").map(i => i.value));
    for (const p of db.produtos) {
      const tem = p.descritivos.includes(d.id), quer = marcados.has(p.id);
      if (tem === quer) continue;
      if (quer) { const { error } = await sb.from("produto_descritivos").insert({ produto_id: p.id, descritivo_id: d.id, ordem: p.descritivos.length }); erroSb(error, "vincular produto"); p.descritivos.push(d.id); }
      else { const { error } = await sb.from("produto_descritivos").delete().eq("produto_id", p.id).eq("descritivo_id", d.id); erroSb(error, "desvincular produto"); p.descritivos = p.descritivos.filter(x => x !== d.id); }
    }
    fecharModal(); renderListaDesc(); renderListaProd(); renderDescs(); toast("Descritivo salvo.");
  });
}
function ligarDescritivos() {
  $("#buscaDesc").addEventListener("input", renderListaDesc);
  $("#btnNovoDesc").addEventListener("click", () => modalDesc(null));
  $("#listaDesc").addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return; const id = b.closest("[data-de]").dataset.de; const d = db.descritivos.find(x => x.id === id);
    if (b.dataset.da === "edit") modalDesc(d);
    if (b.dataset.da === "dup") modalDesc(Object.assign(clone(d), { id: uuid(), titulo: d.titulo + " (cópia)" }), true);
    if (b.dataset.da === "del" && confirm("Excluir o descritivo " + d.titulo + "?")) tentar(async () => {
      await apagar("descritivos", id); db.descritivos = db.descritivos.filter(x => x.id !== id); db.produtos.forEach(p => p.descritivos = p.descritivos.filter(x => x !== id));
      renderListaDesc(); renderDescs(); toast("Descritivo excluído.");
    });
  });
}

/* ============================ configurações ============================ */
function renderConfig() {
  for (const el of $$("[data-cfg]")) { const v = db.config[el.dataset.cfg]; if (el.type === "checkbox") el.checked = !!v; else el.value = v ?? ""; }
  $("#modeloInfo").innerHTML = db.modelo ? `<span style="color:var(--ok)">●</span> Modelo carregado: <b>${esc(db.modeloInfo.nome)}</b>${db.modeloInfo.em ? " (enviado em " + new Date(db.modeloInfo.em).toLocaleString("pt-BR") + ")" : ""}.`
    : `<span style="color:var(--erro)">●</span> Nenhum modelo carregado. Clique em <b>Carregar modelo</b> e escolha o arquivo <code>modelo-v2.docx</code> (pasta <code>gerador-atual</code> no OneDrive).`;
  $("#avisoModelo").innerHTML = db.modelo ? "" : `<div class="aviso">Falta carregar o modelo Word da proposta. <button class="btn peq pri" onclick="irAba('config')">Ir para Configurações</button></div>`;
  $("#dadosInfo").textContent = `${db.clientes.length} clientes · ${db.produtos.length} produtos · ${db.descritivos.length} descritivos · ${db.licencas.length} licenças · ${db.propostas.length} propostas — carregado em ${db.carregadoEm ? db.carregadoEm.toLocaleString("pt-BR") : "—"} · ${VERSAO_APP}`;
}
function ligarConfig() {
  let t = null;
  $$("[data-cfg]").forEach(el => el.addEventListener("change", () => {
    const k = el.dataset.cfg; db.config[k] = el.type === "checkbox" ? el.checked : el.type === "number" ? +el.value : el.value;
    clearTimeout(t); t = setTimeout(() => tentar(async () => { await gravarConfig("proposta_padrao", db.config); toast("Configuração salva."); }), 400);
    if (k === "tituloCapa" && !cur.tituloManual) $("#pTitulo").value = tituloCapa();
    if (k === "arredondamento" || k === "textoCortesia") renderTabelas();
  }));
  $("#btnModeloBaixar").addEventListener("click", () => { if (!db.modelo) return toast("Nenhum modelo carregado.", true); baixarBlob(new Blob([b64ToBuf(db.modelo)]), db.modeloInfo.nome || "modelo-proposta-ska.docx"); });
  $("#btnModeloCarregar").addEventListener("click", () => $("#fileModelo").click());
  $("#fileModelo").addEventListener("change", async e => {
    const f = e.target.files[0]; if (!f) return;
    try {
      const buf = await f.arrayBuffer(); const z = new PizZip(buf);
      const doc = new window.docxtemplater(z, { paragraphLoop: true, linebreaks: true, nullGetter: () => "" }); doc.render(dadosDocumento());
      carregando(true, "Enviando modelo...");
      const b64 = bufToB64(buf);
      await gravarConfig("modelo_docx", { nome: f.name, b64 });
      db.modelo = b64; db.modeloInfo = { nome: f.name, em: new Date().toISOString() }; renderConfig(); toast("Modelo carregado, testado e guardado no banco.");
    } catch (er) { const det = er.properties && er.properties.errors ? er.properties.errors.map(x => x.properties && x.properties.explanation).join("; ") : er.message; toast("Modelo com problema: " + det, true); }
    finally { carregando(false); e.target.value = ""; }
  });
  $("#btnRecarregar").addEventListener("click", () => tentar(async () => { await carregarTudo(); renderTudo(); toast("Dados recarregados."); }));
  $("#btnBackup").addEventListener("click", () => {
    const { modelo, ...resto } = db;
    baixarBlob(new Blob([JSON.stringify(resto, null, 1)], { type: "application/json" }), `backup-propostas-ska-${hojeISO()}.json`);
    toast("Atenção: o backup contém dados pessoais de clientes — guarde em local seguro.");
  });
  $("#btnSair").addEventListener("click", async () => { if (!confirm("Sair do sistema?")) return; await sb.auth.signOut(); location.reload(); });
  $("#usuario").addEventListener("click", () => irAba("config"));
}

/* ============================ geral ============================ */
function abrirModal(html) { $("#modal").style.maxWidth = ""; $("#modal").innerHTML = html; $("#modalBg").classList.add("on"); $("#modalBg").scrollTop = 0; }
function fecharModal() { $("#modalBg").classList.remove("on"); }
window.fecharModal = fecharModal;
$("#modalBg").addEventListener("mousedown", e => { if (e.target.id === "modalBg") fecharModal(); });
document.addEventListener("keydown", e => { if (e.key === "Escape") fecharModal(); });
function irAba(a) {
  $$("#nav button").forEach(b => b.classList.toggle("ativo", b.dataset.aba === a));
  $$(".aba").forEach(s => s.classList.toggle("ativa", s.id === "aba-" + a));
  $("#barraAcao").style.display = a === "proposta" ? "" : "none";
  const b = $(`#nav button[data-aba="${a}"]`); if (b) b.scrollIntoView({ inline: "nearest", block: "nearest" });
  window.scrollTo(0, 0);
}
window.irAba = irAba;
function renderDatalists() {
  $("#dlClientes").innerHTML = db.clientes.filter(c => c.ativo !== false).map(c => `<option value="${esc(rotuloCliente(c))}">`).join("");
  $("#dlProdutos").innerHTML = db.produtos.map(p => `<option value="${esc(p.nome)}">`).join("");
  const unicos = k => [...new Set(db.produtos.map(p => p[k]).filter(Boolean))].sort();
  $("#dlNegocios").innerHTML = unicos("negocio").map(v => `<option value="${esc(v)}">`).join("");
  $("#dlLinhas").innerHTML = unicos("linha").map(v => `<option value="${esc(v)}">`).join("");
  $("#dlSubgrupos").innerHTML = unicos("subgrupo").map(v => `<option value="${esc(v)}">`).join("");
  $("#dlGrupos").innerHTML = db.grupos.map(g => `<option value="${esc(g)}">`).join("");
}
function renderTudo() {
  for (const f of [renderDatalists, preencherFiltrosCli, renderProposta, renderListaProp, renderRenovacoes, renderListaCli, renderListaProd, renderListaDesc, renderConfig]) {
    try { f(); } catch (e) {
      console.error(f.name, e);
      if (f === renderProposta) { cur = novaProposta(); salvarRascunho(); try { renderProposta(); } catch { } mostrarErro("O rascunho da proposta em edição estava inválido e foi descartado (as propostas salvas não foram afetadas).", e); }
      else mostrarErro("Erro ao exibir a seção " + f.name, e);
    }
  }
}
function mostrarErro(msg, e) {
  const box = $("#avisoErro"); if (!box) return;
  box.innerHTML = `<div class="aviso" style="background:#fdecea;border-color:#f5b5ae;color:#8a1c12"><span style="flex:1"><b>Atenção:</b> ${esc(msg)}${e ? `<br><small>Detalhe técnico: ${esc(e.message || e)}</small>` : ""}</span>
    <button class="btn peq" onclick="this.closest('.aviso').remove()">Fechar</button></div>`;
}
window.addEventListener("error", e => mostrarErro("Ocorreu um erro no app. Se algo parar de responder, recarregue a página (F5).", e.error || e.message));
window.addEventListener("unhandledrejection", e => mostrarErro("Ocorreu um erro no app. Se algo parar de responder, recarregue a página (F5).", e.reason));

/* ============================ login e início ============================ */
let iniciado = false;
async function iniciarApp(sessao) {
  usuario = sessao.user;
  $("#login").style.display = "none"; $("#app").style.display = "";
  $("#usuario").textContent = usuario.email;
  if (iniciado) return; iniciado = true;
  try { await carregarTudo(); }
  catch (e) { mostrarErro("Não consegui carregar os dados do banco. Verifique a internet e clique em Configurações → Recarregar dados.", e); }
  try { cur = JSON.parse(localStorage.getItem(LS_RASCUNHO) || "null"); } catch { cur = null; }
  if (!cur || !cur.cliente || !Array.isArray(cur.tabelas)) cur = novaProposta();
  $$("#nav button").forEach(b => b.addEventListener("click", () => irAba(b.dataset.aba)));
  for (const lig of [ligarProposta, ligarListaProp, ligarRenovacoes, ligarClientes, ligarProdutos, ligarDescritivos, ligarConfig]) { try { lig(); } catch (e) { console.error(e); mostrarErro("Erro ao preparar a tela (" + lig.name + ")", e); } }
  $("#btnGerar").addEventListener("click", gerarDocx);
  $("#btnSalvarProp").addEventListener("click", () => salvarProposta(false));
  $("#btnNovaProp").addEventListener("click", () => { if (!confirm("Começar uma proposta nova? (salve antes se quiser manter a atual)")) return; cur = novaProposta(); salvarRascunho(); renderProposta(); });
  renderTudo();
}
function mostrarLogin(msg) { $("#app").style.display = "none"; $("#login").style.display = ""; $("#formLogin").style.display = ""; $("#formNovaSenha").style.display = "none"; if (msg) $("#lgMsg").textContent = msg; }
$("#formLogin").addEventListener("submit", async e => {
  e.preventDefault(); const b = $("#lgEntrar"); b.disabled = true; b.innerHTML = '<span class="spin"></span> Entrando'; $("#lgMsg").textContent = "";
  const { data, error } = await sb.auth.signInWithPassword({ email: $("#lgEmail").value.trim(), password: $("#lgSenha").value });
  b.disabled = false; b.textContent = "Entrar";
  if (error) { $("#lgMsg").innerHTML = `<span style="color:var(--erro)">${/invalid/i.test(error.message) ? "E-mail ou senha incorretos." : /confirm/i.test(error.message) ? "E-mail ainda não confirmado — veja sua caixa de entrada." : esc(error.message)}</span>`; return; }
  iniciarApp(data.session);
});
$("#lgEsqueci").addEventListener("click", async () => {
  const email = $("#lgEmail").value.trim(); if (!email) { $("#lgMsg").textContent = "Digite o seu e-mail acima e clique de novo em \"Esqueci a senha\"."; return; }
  const { error } = await sb.auth.resetPasswordForEmail(email, { redirectTo: location.origin + location.pathname });
  $("#lgMsg").textContent = error ? "Erro: " + error.message : "Se o e-mail estiver cadastrado, você vai receber um link para criar uma nova senha.";
});
$("#formNovaSenha").addEventListener("submit", async e => {
  e.preventDefault();
  const { error } = await sb.auth.updateUser({ password: $("#nsSenha").value });
  if (error) { $("#nsMsg").textContent = "Erro: " + error.message; return; }
  $("#nsMsg").textContent = "Senha alterada."; history.replaceState(null, "", location.pathname);
  const { data } = await sb.auth.getSession(); if (data.session) iniciarApp(data.session);
});
sb.auth.onAuthStateChange((ev, sessao) => {
  if (ev === "PASSWORD_RECOVERY") { $("#app").style.display = "none"; $("#login").style.display = ""; $("#formLogin").style.display = "none"; $("#formNovaSenha").style.display = ""; return; }
  if (ev === "SIGNED_OUT") mostrarLogin();
});
(async () => {
  const { data } = await sb.auth.getSession();
  if (data.session && !/type=recovery/.test(location.hash)) iniciarApp(data.session); else if (!/type=recovery/.test(location.hash)) mostrarLogin();
})();
