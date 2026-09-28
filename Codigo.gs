/**
 * =============================================================================
 * CALCULADORA DE PREÇO · EXEMPLO (v1.1.0)
 * -----------------------------------------------------------------------------
 * Projeto Apps Script SEPARADO do Monitor de Honorários.
 *
 * Por que separado: tudo que um projeto do Apps Script tem de função pode ser
 * chamado por quem abre o link dele. Esconder a tela financeira dentro do
 * Monitor não protegeria o dado. Aqui o projeto simplesmente não tem acesso
 * ao financeiro: ele só lê a planilha "Calculadora de Preço - Base", que o
 * Monitor alimenta com o mínimo necessário para calcular:
 *   CNPJ, nome, código Domínio, atividade, regime, faturamento de 12 meses,
 *   média de notas, funcionários, soma da raiz do CNPJ, pendência fiscal.
 * Nenhum honorário, recebimento, inadimplência ou resultado da carteira.
 *
 * Abas da planilha "Calculadora de Preço - Base":
 *   CALC_EMPRESAS      escrita pelo Monitor (publicarBaseCalculadora)
 *   CALC_CONFIG        salário mínimo e data da última atualização
 *   PRECIF_DIRETRIZES  regras da tabela (a mesma que o Monitor usa)
 *   PRECIF_TESTES      testes com clientes
 *   PRECIF_EDITORES    quem pode mudar as regras (vazia = todos)
 *
 * Implantação: App da Web · Executar como: EU (dono da planilha) ·
 * Quem pode acessar: qualquer pessoa da Exemplo. Assim a equipe usa a
 * calculadora sem precisar de acesso à planilha.
 * A ID da planilha já vem gravada em CALC_PLANILHA_PADRAO. A propriedade
 * CALC_PLANILHA_ID só é necessária se um dia a planilha mudar.
 * =============================================================================
 */

/* Planilha "Calculadora de Preço - Base" (Drive da Exemplo) */
var CALC_PLANILHA_PADRAO = 'ID_EXEMPLO';

var CALC_ABA_EMPRESAS = 'CALC_EMPRESAS';
var CALC_ABA_CONFIG   = 'CALC_CONFIG';
var PRECIF_ABA_DIRETRIZES = 'PRECIF_DIRETRIZES';
var PRECIF_ABA_TESTES     = 'PRECIF_TESTES';
var PRECIF_ABA_EDITORES   = 'PRECIF_EDITORES';
var PRECIF_CAB_DIRETRIZES = ['chave', 'valor', 'status', 'fonte', 'criterio', 'alterado_por', 'alterado_em'];
var PRECIF_CAB_TESTES     = ['id', 'data', 'cnpj', 'empresa', 'honorario_atual', 'sugerido', 'diferenca',
                             'situacao', 'usuario', 'observacao', 'entrada_json'];
var PRECIF_CAB_EDITORES   = ['email', 'nome', 'observacao'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Calculadora de preço · Exemplo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/* ------------------------------------------------------------ utilitários */
function calc_planilha_() {
  var id = PropertiesService.getScriptProperties().getProperty('CALC_PLANILHA_ID') || CALC_PLANILHA_PADRAO;
  return SpreadsheetApp.openById(id);
}
function calc_aba_(nome, cabecalho) {
  var ss = calc_planilha_();
  var aba = ss.getSheetByName(nome);
  if (!aba && cabecalho) {
    aba = ss.insertSheet(nome);
    aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}
function calc_linhas_(aba) {
  if (!aba) return [];
  var ult = aba.getLastRow();
  if (ult < 2) return [];
  return aba.getRange(2, 1, ult - 1, aba.getLastColumn()).getValues();
}
function calc_email_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}
function calc_podeEditar_() {
  var lista = calc_linhas_(calc_aba_(PRECIF_ABA_EDITORES, PRECIF_CAB_EDITORES))
    .map(function (l) { return String(l[0] || '').trim().toLowerCase(); })
    .filter(function (e) { return e; });
  if (!lista.length) return true;
  return lista.indexOf(calc_email_().toLowerCase()) >= 0;
}
function calc_iso_(v) { return v instanceof Date ? v.toISOString() : (v ? String(v) : ''); }
function calc_num_(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(v); return isNaN(n) ? null : n;
}

/** Rodar uma vez pelo editor. */
function configurarCalculadora() {
  calc_aba_(PRECIF_ABA_DIRETRIZES, PRECIF_CAB_DIRETRIZES);
  calc_aba_(PRECIF_ABA_TESTES, PRECIF_CAB_TESTES);
  calc_aba_(PRECIF_ABA_EDITORES, PRECIF_CAB_EDITORES);
  return 'Abas prontas. Agora rode publicarBaseCalculadora() no projeto do Monitor.';
}

/* ================================================================== API */

/** Cadastro reduzido para a calculadora. Não existe dado financeiro aqui. */
function api_calcBase() {
  var cfg = {};
  calc_linhas_(calc_aba_(CALC_ABA_CONFIG)).forEach(function (l) { cfg[String(l[0])] = l[1]; });
  var empresas = calc_linhas_(calc_aba_(CALC_ABA_EMPRESAS)).map(function (l) {
    return { cnpj: String(l[0] || '').replace(/\D/g, ''), empresa: l[1] || '', codigoDominio: String(l[2] || ''),
             ativ: l[3] || 'servico', regime: l[4] || 'SN', rbt12: Number(l[5]) || 0, notasMes: Number(l[6]) || 0,
             funcionarios: Number(l[7]) || 0, outrosQtd: Number(l[8]) || 0, outrosRbt: Number(l[9]) || 0,
             pendenciaFiscal: String(l[10]) === 'Sim' };
  }).filter(function (e) { return e.cnpj; });
  return { ok: true, empresas: empresas,
           salarioMinimo: Number(cfg.salario_minimo) || 1621,
           atualizadoEm: calc_iso_(cfg.atualizado_em), usuario: calc_email_() };
}

function api_precCarregar() {
  var itens = {};
  calc_linhas_(calc_aba_(PRECIF_ABA_DIRETRIZES, PRECIF_CAB_DIRETRIZES)).forEach(function (l) {
    var ch = String(l[0] || '').trim();
    if (!ch) return;
    itens[ch] = { valor: calc_num_(l[1]), status: l[2] || 'Provisório', fonte: l[3] || '',
                  obs: l[4] || '', alteradoPor: l[5] || '', alteradoEm: calc_iso_(l[6]) };
  });
  var testes = calc_linhas_(calc_aba_(PRECIF_ABA_TESTES, PRECIF_CAB_TESTES)).map(function (l) {
    var entrada = {};
    try { entrada = JSON.parse(l[10] || '{}'); } catch (e) { entrada = {}; }
    /* Testes feitos no Monitor ficam na mesma aba e podem ter o honorário
       atual. Aqui ele nunca sai: nem valor, nem diferença, nem situação. */
    delete entrada.atual; delete entrada.inad;
    return { id: String(l[0]), data: calc_iso_(l[1]), cnpj: String(l[2] || ''), empresa: l[3] || '',
             atual: null, sugerido: calc_num_(l[5]), diferenca: null,
             situacao: l[7] === 'Reunião' ? 'Reunião' : 'Proposta', usuario: l[8] || '',
             observacao: l[9] || '', entrada: entrada };
  }).filter(function (t) { return t.id; });
  return { ok: true, diretrizes: itens, testes: testes, podeEditar: calc_podeEditar_(), usuario: calc_email_() };
}

function api_precSalvarDiretrizes(lista, usuario) {
  if (!calc_podeEditar_()) return { ok: false, erro: 'Seu acesso à calculadora é só de consulta.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var aba = calc_aba_(PRECIF_ABA_DIRETRIZES, PRECIF_CAB_DIRETRIZES);
    var posicao = {};
    calc_linhas_(aba).forEach(function (l, i) { posicao[String(l[0])] = i + 2; });
    var quem = calc_email_() || usuario || 'equipe Exemplo';
    var agora = new Date();
    (lista || []).forEach(function (m) {
      var linha = [m.chave, (m.valor === '' || m.valor === null) ? '' : Number(m.valor),
                   m.status || 'Provisório', m.fonte || '', m.obs || '', quem, agora];
      if (posicao[m.chave]) aba.getRange(posicao[m.chave], 1, 1, linha.length).setValues([linha]);
      else { aba.appendRow(linha); posicao[m.chave] = aba.getLastRow(); }
    });
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return { ok: true, diretrizes: api_precCarregar().diretrizes, mensagem: (lista || []).length + ' regra(s) salva(s).' };
}

function api_precSalvarTeste(t) {
  t = t || {};
  var id = 't' + new Date().getTime(), agora = new Date();
  var quem = calc_email_() || t.usuario || '';
  calc_aba_(PRECIF_ABA_TESTES, PRECIF_CAB_TESTES).appendRow([
    id, agora, t.cnpj ? "'" + t.cnpj : '', t.empresa || '',
    '', t.sugerido === null || t.sugerido === undefined ? '' : t.sugerido, '',
    t.situacao === 'Reunião' ? 'Reunião' : 'Proposta', quem, t.observacao || '',
    JSON.stringify(calc_semFinanceiro_(t.entrada))
  ]);
  var reg = {};
  for (var k in t) reg[k] = t[k];
  reg.id = id; reg.data = agora.toISOString(); reg.usuario = quem;
  reg.atual = null; reg.diferenca = null; reg.entrada = calc_semFinanceiro_(t.entrada);
  reg.situacao = t.situacao === 'Reunião' ? 'Reunião' : 'Proposta';
  return { ok: true, teste: reg, mensagem: 'Teste salvo.' };
}

function calc_semFinanceiro_(e) {
  var o = {};
  for (var k in (e || {})) if (k !== 'atual' && k !== 'inad') o[k] = e[k];
  return o;
}

function api_precExcluirTeste(id) {
  var aba = calc_aba_(PRECIF_ABA_TESTES, PRECIF_CAB_TESTES);
  var linhas = calc_linhas_(aba);
  for (var i = linhas.length - 1; i >= 0; i--) {
    if (String(linhas[i][0]) === String(id)) { aba.deleteRow(i + 2); return { ok: true }; }
  }
  return { ok: false, erro: 'Teste não encontrado.' };
}
