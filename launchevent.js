/*
 * Szablon spotkania – Outlook add-in
 * Automatycznie wstawia agendę i link do logowania czasu
 * do opisu każdego nowego spotkania (OnNewAppointmentOrganizer).
 *
 * >>> KONFIGURACJA – zmień tylko ten blok <<<
 */
var CONFIG = {
  // Napis pod agendą – użytkownik sam wkleja link po dwukropku
  timeLogLabel: "Link do logowania czasu",
  // Liczba pustych punktów agendy
  agendaItems: 3,
  // Unikalny znacznik – po nim add-in rozpoznaje, że szablon już jest w opisie
  marker: "Agenda:"
};

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildHtml() {
  var items = "";
  for (var i = 0; i < CONFIG.agendaItems; i++) {
    items += "<li>&nbsp;</li>";
  }
  return (
    "<p><b>" + escapeHtml(CONFIG.marker) + "</b></p>" +
    "<ol>" + items + "</ol>" +
    "<p>" + escapeHtml(CONFIG.timeLogLabel) + ":&nbsp;</p>" +
    "<p>&nbsp;</p>"
  );
}

function buildText() {
  var lines = [CONFIG.marker];
  for (var i = 1; i <= CONFIG.agendaItems; i++) {
    lines.push(i + ". ");
  }
  lines.push("");
  lines.push(CONFIG.timeLogLabel + ": ");
  lines.push("");
  return lines.join("\n");
}

/**
 * Wstawia szablon na początek opisu, jeśli jeszcze go tam nie ma.
 * done() jest wywoływane zawsze, także przy błędzie.
 */
function insertTemplate(done) {
  var item = Office.context.mailbox.item;

  item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
    var current = bodyResult.status === Office.AsyncResultStatus.Succeeded ? bodyResult.value || "" : "";
    if (current.indexOf(CONFIG.marker) !== -1) {
      done();
      return;
    }

    item.body.getTypeAsync(function (typeResult) {
      var isHtml =
        typeResult.status !== Office.AsyncResultStatus.Succeeded ||
        typeResult.value === Office.CoercionType.Html;

      item.body.prependAsync(
        isHtml ? buildHtml() : buildText(),
        { coercionType: isHtml ? Office.CoercionType.Html : Office.CoercionType.Text },
        function (result) {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Szablon spotkania: " + JSON.stringify(result.error));
          }
          done();
        }
      );
    });
  });
}

// Wywoływane automatycznie przy tworzeniu nowego spotkania
function onNewAppointmentHandler(event) {
  insertTemplate(function () {
    event.completed();
  });
}

// Przycisk „Wstaw szablon” na wstążce (np. gdy ktoś usunie szablon)
function insertTemplateCommand(event) {
  insertTemplate(function () {
    event.completed();
  });
}

Office.onReady(function () {});

if (Office.actions && Office.actions.associate) {
  Office.actions.associate("onNewAppointmentHandler", onNewAppointmentHandler);
  Office.actions.associate("insertTemplateCommand", insertTemplateCommand);
}
