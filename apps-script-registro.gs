// Cole este código em script.google.com (ou em uma planilha, em Extensões > Apps Script).
// Ele recebe o nome e o Gmail de quem entra no app e envia um e-mail para você.

var TOKEN = 'troque-por-uma-senha-sua';          // tem que ser igual ao token do config.js
var ADMIN_EMAIL = 'leo.miserer123@gmail.com';

function doPost(e) {
  var d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) {}
  if (d.token !== TOKEN) return ContentService.createTextOutput('negado');

  var name = String(d.name || '').replace(/[\r\n]/g, ' ').slice(0, 80);
  var email = String(d.email || '').replace(/[\r\n]/g, ' ').slice(0, 120);
  var when = String(d.at || new Date().toISOString());
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return ContentService.createTextOutput('invalido');

  // Registra em planilha, se o script estiver ligado a uma.
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.getSheets()[0].appendRow([when, name, email]);
  } catch (err) {}

  MailApp.sendEmail(
    ADMIN_EMAIL,
    'Novo acesso no Estúdio UGC: ' + name,
    'Nome: ' + name + '\nGmail: ' + email + '\nHorário: ' + when
  );
  return ContentService.createTextOutput('ok');
}
