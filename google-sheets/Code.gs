/**
 * TraffiCount — recolha das contagens numa Google Sheet.
 * Powered by GISON – Soluções em Geotecnologias, Lda.
 *
 * Instalação (uma vez):
 *  1. Crie uma Google Sheet nova no seu Drive (ex.: «TraffiCount – Contagens»).
 *  2. Menu Extensões → Apps Script. Apague o conteúdo e cole este ficheiro todo.
 *  3. Altere o CODIGO_EQUIPA abaixo para um código só vosso (ex.: PRN-2026-xyz).
 *  4. Implementar → Nova implementação → tipo «Aplicação Web».
 *       Executar como: Eu  ·  Quem tem acesso: Qualquer pessoa
 *     Autorize o acesso quando o Google pedir e copie o URL da aplicação web (termina em /exec).
 *  5. Na aplicação TraffiCount: Contagens → Nuvem → Configurar, cole o URL e o código.
 *
 * As folhas «Sessoes» e «Eventos» são criadas automaticamente no primeiro envio.
 */

const CODIGO_EQUIPA = 'ALTERAR-ESTE-CODIGO';

const CABECALHO_SESSOES = [
  'sessao_id', 'data', 'inicio', 'fim', 'duracao_min', 'n_pausas', 'observador',
  'sentido1_chegada', 'sentido2_saida', 'lat', 'lon', 'precisao_m', 'origem_posicao',
  'MC_cheg', 'MC_saida', 'LP_cheg', 'LP_saida', 'LM_cheg', 'LM_saida', 'PP_cheg', 'PP_saida', 'PM_cheg', 'PM_saida',
  'total_cheg', 'total_saida', 'total', 'pesados_pct', 'veic_hora', 'versao_app', 'mapa', 'recebido_em'
];
const CABECALHO_EVENTOS = ['sessao_id', 't_ms', 'data_hora', 'tipo', 'sentido', 'categoria'];

function doGet() {
  return resposta({ ok: true, app: 'TraffiCount', folha: SpreadsheetApp.getActive().getName() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const pedido = JSON.parse(e.postData.contents);
    if (pedido.codigo !== CODIGO_EQUIPA) return resposta({ ok: false, erro: 'codigo' });
    if (pedido.teste) return resposta({ ok: true, app: 'TraffiCount', folha: SpreadsheetApp.getActive().getName() });

    const ss = SpreadsheetApp.getActive();
    const folhaS = obterFolha(ss, 'Sessoes', CABECALHO_SESSOES);
    const folhaE = obterFolha(ss, 'Eventos', CABECALHO_EVENTOS);
    const ids = folhaS.getLastRow() > 1
      ? folhaS.getRange(2, 1, folhaS.getLastRow() - 1, 1).getValues().map(r => String(r[0]))
      : [];

    const resultados = (pedido.sessoes || []).map(s => {
      if (ids.indexOf(s.id) >= 0) return { id: s.id, estado: 'ja_existia' };
      // envio anterior interrompido a meio: remove eventos órfãos desta sessão antes de os gravar de novo
      apagarEventos(folhaE, s.id);
      if (s.eventos && s.eventos.length) {
        folhaE.getRange(folhaE.getLastRow() + 1, 1, s.eventos.length, CABECALHO_EVENTOS.length).setValues(s.eventos);
      }
      // a linha da sessão é gravada por último: só existe quando os eventos estão completos
      folhaS.appendRow(s.linha.concat([new Date()]));
      ids.push(s.id);
      return { id: s.id, estado: 'gravada' };
    });
    return resposta({ ok: true, resultados: resultados });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function obterFolha(ss, nome, cabecalho) {
  let f = ss.getSheetByName(nome);
  if (!f) {
    f = ss.insertSheet(nome);
    f.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight('bold');
    f.setFrozenRows(1);
  }
  return f;
}

function apagarEventos(folha, id) {
  if (folha.getLastRow() < 2) return;
  const linhas = folha.getRange(2, 1, folha.getLastRow() - 1, 1)
    .createTextFinder(id).matchEntireCell(true).findAll()
    .map(r => r.getRow()).sort((a, b) => b - a);
  linhas.forEach(r => folha.deleteRow(r));
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
