const SHEET_NAME = 'SAC';
const SAC_EMAIL = 'sac@vegasvigilancia.com.br';
const SHEET_ID = '1yOBD4WZRWWDHCfpPeUOy7iKuIoijfv2ZNhRSZLRqozo';
const LOGO_URL = 'https://vegasvig.github.io/SAC/assets/logovegas.png';
const POSVENDA_EMAIL = 'gra.posvenda@vegasvigilancia.com.br';
const PAINEL_URL = 'https://vegasvig.github.io/SAC/';
const PRAZO_RETIRADA = 30;
const COL_RETIRADA = 14; // N

function doPost(e) {
  const p = e.parameter;
  const r = { ok: true, sheet: false, mail: false };
  try {
    const status = (p.decisao || '').toLowerCase().includes('proposta')
      ? 'Quer Proposta' : 'Cancelamento';
    SpreadsheetApp.openById(SHEET_ID).getSheets()[0].appendRow([
      new Date(),    // A - Data/Hora
      p.protocolo,   // B - Protocolo
      p.nome,        // C - Nome
      p.cpf,         // D - Cpf/Cnpj
      p.telefone,    // E - Telefone
      p.email,       // F - E-mail
      p.nota,        // G - Satisfação
      p.motivos,     // H - Motivos
      p.descricao,   // I - Descrição
      p.decisao,     // J - Decisão
      status,        // K - Status
      '',            // L - Respondido
      p.NPS,         // M - NPS
      ''             // N - Retirada
    ]);
    r.sheet = true;
  } catch (err) { r.erroSheet = String(err); }

  if (p.enviar_email === 'sim') {
    try { enviarEmails(p); r.mail = true; }
    catch (err) { r.erroMail = String(err); }
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function enviarEmails(p) {
  const logoBlob = UrlFetchApp.fetch(LOGO_URL).getBlob().setName('logovegas.png');
  const opts = {
    htmlBody: montarHtmlEmail(p),
    name: 'Vegas Vigilância e Segurança',
    inlineImages: { logo: logoBlob }
  };
  MailApp.sendEmail(SAC_EMAIL, `Solicitação de Cancelamento — ${p.protocolo}`, 'Nova solicitação registrada.', opts);
  if (p.email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p.email)) {
    MailApp.sendEmail(p.email, `Confirmação de Solicitação — ${p.protocolo}`, 'Sua solicitação foi registrada.', opts);
  }
}

function montarHtmlEmail(p) {
  const row = (l, v) => `<tr>
    <td style="padding:8px 12px;background:#141414;color:#888;font-size:12px;border:1px solid #2a2a2a;width:38%">${l}</td>
    <td style="padding:8px 12px;background:#0f0f0f;color:#f0f0f0;font-size:13px;border:1px solid #2a2a2a">${v || 'Não informado'}</td>
  </tr>`;
  return `
  <div style="max-width:600px;margin:auto;font-family:Arial,sans-serif;background:#0a0a0a;padding:24px;border-radius:8px">
    <div style="text-align:center;padding-bottom:16px;border-bottom:1px solid #2a2a2a;margin-bottom:16px">
      <img src="cid:logo" alt="Vegas Vigilância e Segurança" width="200" style="max-width:200px;height:auto;display:inline-block">
    </div>
    <h2 style="color:#c0c0c0;border-bottom:1px solid #2a2a2a;padding-bottom:12px;margin:0 0 16px">
      Solicitação de Cancelamento Registrada
    </h2>
    <p style="color:#aaa;font-size:13px">Protocolo <strong style="color:#fff">${p.protocolo}</strong> — ${p.data_hora}</p>
    <table style="width:100%;border-collapse:collapse;margin-top:12px">
      ${row('Nome', p.nome)}
      ${row('CPF / CNPJ', p.cpf)}
      ${row('Telefone', p.telefone)}
      ${row('E-mail', p.email)}
      ${row('Satisfação', p.nota)}
      ${row('Motivos', p.motivos)}
      ${row('Descrição', p.descricao)}
      ${row('Decisão', p.decisao)}
    </table>
    <p style="color:#666;font-size:11px;margin-top:20px;border-top:1px solid #2a2a2a;padding-top:12px">
      Nossa equipe entrará em contato em até 48 horas úteis.<br>
      © 2026 Vegas Vigilância • CNPJ 29.890.721/0001-04
    </p>
  </div>`;
}

function doGet(e) {
  const cb = e.parameter.callback || '';
  const action = e.parameter.action || '';

  if (action === 'marcarRespondido') {
    const protocolo = e.parameter.protocolo || '';
    try {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
      const rows = sheet.getDataRange().getValues();

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][1]) === protocolo) {
          sheet.getRange(i + 1, 12).setValue('Respondido');
          sheet.getRange(i + 1, 1, 1, sheet.getLastColumn())
            .setBackground('#1a3a1a')
            .setFontColor('#ccffcc');
          break;
        }
      }

      const res = JSON.stringify({ ok: true });
      return ContentService
        .createTextOutput(cb ? cb + '(' + res + ')' : res)
        .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);

    } catch (err) {
      const res = JSON.stringify({ ok: false, erro: String(err) });
      return ContentService
        .createTextOutput(cb ? cb + '(' + res + ')' : res)
        .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
    }
  }

  if (action === 'marcarRetirada') {
    return responder(cb, marcarRetirada(e.parameter.protocolo || ''));
  }

  const sheet = abaSAC();
  const rows = sheet.getDataRange().getValues();

  const chavesFixas = [
    'data_hora',   // A
    'protocolo',   // B
    'nome',        // C
    'cpf',         // D
    'telefone',    // E
    'email',       // F
    'nota',        // G
    'motivos',     // H
    'descricao',   // I
    'decisao',     // J
    'status',      // K
    'respondido',  // L
    'NPS',         // M
    'retirada'     // N
  ];

  const data = rows.slice(1).map(r => {
    const obj = {};
    chavesFixas.forEach((chave, i) => {
      obj[chave] = r[i] !== undefined ? r[i] : '';
    });
    return obj;
  });

  const json = JSON.stringify({ data });
  const output = cb ? cb + '(' + json + ')' : json;
  const mime = cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(output).setMimeType(mime);
}


function abaSAC() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function responder(cb, obj) {
  const json = JSON.stringify(obj);
  return ContentService
    .createTextOutput(cb ? cb + '(' + json + ')' : json)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function marcarRetirada(protocolo) {
  if (!protocolo) return { ok: false, erro: 'Protocolo vazio' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = abaSAC();
    const protos = sheet.getRange(2, 2, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
    for (let i = 0; i < protos.length; i++) {
      if (String(protos[i][0]) === protocolo) {
        sheet.getRange(i + 2, COL_RETIRADA).setValue(new Date());
        return { ok: true };
      }
    }
    return { ok: false, erro: 'Protocolo não encontrado' };
  } catch (err) {
    return { ok: false, erro: String(err) };
  } finally {
    lock.releaseLock();
  }
}

function diasRestantesRetirada(dataSolicitacao, hoje) {
  const ini = new Date(dataSolicitacao);
  ini.setHours(0, 0, 0, 0);
  const passados = Math.round((hoje - ini) / 86400000);
  return PRAZO_RETIRADA - passados;
}

function faixaRetirada(dias) {
  if (dias < 0) return 'vencido';
  if (dias <= 7) return 'vermelho';
  if (dias <= 15) return 'amarelo';
  return 'verde';
}

function listarRetiradasPendentes() {
  const rows = abaSAC().getDataRange().getValues().slice(1);
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  return rows
    .filter(r => r[0] && String(r[10]) === 'Cancelamento' && !r[COL_RETIRADA - 1])
    .map(r => {
      const dias = diasRestantesRetirada(r[0], hoje);
      const prazo = new Date(r[0]);
      prazo.setDate(prazo.getDate() + PRAZO_RETIRADA);
      return {
        protocolo: r[1], nome: r[2], telefone: r[4],
        solicitado: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        prazo: Utilities.formatDate(prazo, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        dias, faixa: faixaRetirada(dias)
      };
    })
    .sort((a, b) => a.dias - b.dias);
}

function alertaRetiradas() {
  const lista = listarRetiradasPendentes();
  if (!lista.length) return;

  const cfg = {
    vencido:  { titulo: 'Prazo vencido',        cor: '#ff2d55', bg: '#2a0a12' },
    vermelho: { titulo: 'Crítico — 7 a 0 dias', cor: '#ff2d55', bg: '#1f0a10' },
    amarelo:  { titulo: 'Atenção — 15 a 8 dias', cor: '#ffb300', bg: '#1f1700' },
    verde:    { titulo: 'No prazo — 30 a 16 dias', cor: '#00c98a', bg: '#06190f' }
  };
  const esc = v => String(v == null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const cont = { vencido: 0, vermelho: 0, amarelo: 0, verde: 0 };
  lista.forEach(c => cont[c.faixa]++);

  const blocos = ['vencido', 'vermelho', 'amarelo', 'verde']
    .filter(f => cont[f])
    .map(f => {
      const c = cfg[f];
      const linhas = lista.filter(x => x.faixa === f).map(x => `<tr>
        <td style="padding:7px 10px;border:1px solid #2a2a2a;color:#f0f0f0;font-size:12px">${esc(x.protocolo)}</td>
        <td style="padding:7px 10px;border:1px solid #2a2a2a;color:#f0f0f0;font-size:12px">${esc(x.nome)}</td>
        <td style="padding:7px 10px;border:1px solid #2a2a2a;color:#f0f0f0;font-size:12px">${esc(x.telefone)}</td>
        <td style="padding:7px 10px;border:1px solid #2a2a2a;color:#aaa;font-size:12px">${x.solicitado}</td>
        <td style="padding:7px 10px;border:1px solid #2a2a2a;color:#aaa;font-size:12px">${x.prazo}</td>
        <td style="padding:7px 10px;border:1px solid #2a2a2a;color:${c.cor};font-size:12px;font-weight:bold;text-align:center">${x.dias < 0 ? Math.abs(x.dias) + 'd atrasado' : x.dias + 'd'}</td>
      </tr>`).join('');
      return `<div style="margin-top:18px;border-left:4px solid ${c.cor};background:${c.bg};padding:10px 12px;border-radius:4px">
        <div style="color:${c.cor};font-weight:bold;font-size:14px;margin-bottom:8px">${c.titulo} (${cont[f]})</div>
        <table style="width:100%;border-collapse:collapse">
          <tr>${['Protocolo', 'Cliente', 'Telefone', 'Solicitado', 'Prazo final', 'Restam'].map(h => `<th style="padding:6px 10px;border:1px solid #2a2a2a;color:#888;font-size:11px;text-align:left">${h}</th>`).join('')}</tr>
          ${linhas}
        </table>
      </div>`;
    }).join('');

  const urgentes = cont.vencido + cont.vermelho;
  const assunto = urgentes
    ? `⚠️ Retiradas: ${urgentes} crítica(s)/vencida(s) — ${lista.length} pendente(s)`
    : `Retiradas pendentes: ${lista.length} cliente(s)`;

  const html = `<div style="max-width:720px;margin:auto;font-family:Arial,sans-serif;background:#0a0a0a;padding:24px;border-radius:8px">
    <div style="text-align:center;padding-bottom:16px;border-bottom:1px solid #2a2a2a">
      <img src="cid:logo" alt="Vegas Vigilância e Segurança" width="180" style="max-width:180px;height:auto">
    </div>
    <h2 style="color:#c0c0c0;margin:16px 0 4px">Controle de retirada de equipamentos</h2>
    <p style="color:#888;font-size:13px;margin:0">Prazo de ${PRAZO_RETIRADA} dias a partir da solicitação de cancelamento.</p>
    ${blocos}
    <p style="margin-top:20px"><a href="${PAINEL_URL}" style="color:#00d4ff;font-size:13px">Abrir painel do SAC</a></p>
    <p style="color:#666;font-size:11px;margin-top:16px;border-top:1px solid #2a2a2a;padding-top:12px">Alerta automático diário • Vegas Vigilância</p>
  </div>`;

  const logo = UrlFetchApp.fetch(LOGO_URL).getBlob().setName('logovegas.png');
  MailApp.sendEmail(POSVENDA_EMAIL, assunto, `${lista.length} retirada(s) pendente(s). Abra o painel: ${PAINEL_URL}`, {
    htmlBody: html,
    name: 'Vegas Vigilância e Segurança',
    inlineImages: { logo }
  });
}

function instalarGatilhoRetiradas() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'alertaRetiradas')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alertaRetiradas').timeBased().everyDays(1).atHour(8).create();
}
