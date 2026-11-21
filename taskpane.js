/* Libreria LAMBDA - componente aggiuntivo personale per Excel
   Le formule sono salvate sul computer (localStorage del componente aggiuntivo).
   sintassi "it": formula nella lingua di Excel (SE, ARROTONDA, separatore ;) -> addFormulaLocal
   sintassi "en": formula in inglese (IF, ROUND, separatore ,), come la restituisce Excel -> add
*/
(function () {
  "use strict";

  var CHIAVE = "libreriaLambda.v1";
  var inExcel = false;

  var PREDEFINITE = [
    {
      nome: "IRPEF2026",
      descrizione: "IRPEF lorda 2026: 23% fino a 28.000, 33% fino a 50.000, 43% oltre (art. 11 TUIR, L. 199/2025). Esclude detrazioni e addizionali.",
      sintassi: "it",
      formula: "=LAMBDA(reddito_imponibile; LET(imponibile; SE(reddito_imponibile>0; reddito_imponibile; 0); soglia_1; 28000; soglia_2; 50000; aliq_1; 0,23; aliq_2; 0,33; aliq_3; 0,43; quota_1; SE(imponibile>soglia_1; soglia_1; imponibile); quota_2; SE(imponibile>soglia_2; soglia_2-soglia_1; SE(imponibile>soglia_1; imponibile-soglia_1; 0)); quota_3; SE(imponibile>soglia_2; imponibile-soglia_2; 0); ARROTONDA(quota_1*aliq_1 + quota_2*aliq_2 + quota_3*aliq_3; 2)))"
    },
    {
      nome: "IRPEF2025",
      descrizione: "IRPEF lorda 2025: 23% fino a 28.000, 35% fino a 50.000, 43% oltre (art. 11 TUIR). Esclude detrazioni e addizionali.",
      sintassi: "it",
      formula: "=LAMBDA(reddito_imponibile; LET(imponibile; SE(reddito_imponibile>0; reddito_imponibile; 0); soglia_1; 28000; soglia_2; 50000; aliq_1; 0,23; aliq_2; 0,35; aliq_3; 0,43; quota_1; SE(imponibile>soglia_1; soglia_1; imponibile); quota_2; SE(imponibile>soglia_2; soglia_2-soglia_1; SE(imponibile>soglia_1; imponibile-soglia_1; 0)); quota_3; SE(imponibile>soglia_2; imponibile-soglia_2; 0); ARROTONDA(quota_1*aliq_1 + quota_2*aliq_2 + quota_3*aliq_3; 2)))"
    },
    {
      nome: "CON_IVA",
      descrizione: "Importo maggiorato dell'IVA al 22%.",
      sintassi: "it",
      formula: "=LAMBDA(importo; importo*1,22)"
    },
    {
      nome: "SCORPORO_IVA",
      descrizione: "Imponibile ricavato da un importo comprensivo di IVA al 22%, arrotondato al centesimo.",
      sintassi: "it",
      formula: "=LAMBDA(importo_lordo; ARROTONDA(importo_lordo/1,22; 2))"
    }
  ];

  var libreria = [];
  var inModifica = null;      // nome della voce in modifica
  var daConfermare = {};      // chiave azione -> true (conferma con secondo clic)

  function $(id) { return document.getElementById(id); }

  var ICONE = {
    inserisci: '<path d="M10 3v10M6 9l4 4 4-4"/><path d="M4 16h12"/>',
    modifica: '<path d="M12.5 4.5l3 3L8 15H5v-3z"/><path d="M11 6l3 3"/>',
    elimina: '<path d="M4 6h12M8 6V4h4v2M6 6l1 10h6l1-10"/>',
    freccia: '<path d="M7 5l5 5-5 5"/>',
    ok: '<circle cx="10" cy="10" r="7"/><path d="M7 10.2l2 2 4-4.4"/>',
    errore: '<circle cx="10" cy="10" r="7"/><path d="M10 6.5v4.5M10 13.5v.01"/>',
    info: '<circle cx="10" cy="10" r="7"/><path d="M10 9v4.5M10 6.5v.01"/>'
  };
  function svg(nome) { return '<svg viewBox="0 0 20 20" aria-hidden="true">' + ICONE[nome] + '</svg>'; }

  // ---------- archivio ----------
  function carica() {
    try {
      var dati = window.localStorage.getItem(CHIAVE);
      if (dati) {
        var l = JSON.parse(dati);
        if (Array.isArray(l)) { libreria = l; return; }
      }
    } catch (e) { /* archivio non disponibile */ }
    libreria = PREDEFINITE.slice();
    salva();
  }
  function salva() {
    try { window.localStorage.setItem(CHIAVE, JSON.stringify(libreria)); }
    catch (e) { messaggio("Non riesco a salvare sul computer: fai un backup prima di chiudere.", "errore"); }
  }
  function trova(nome) {
    var n = nome.toUpperCase();
    for (var i = 0; i < libreria.length; i++) { if (libreria[i].nome.toUpperCase() === n) return i; }
    return -1;
  }

  // ---------- interfaccia ----------
  var timerStato = null;
  function messaggio(testo, tipo) {
    var s = $("stato");
    tipo = tipo || "info";
    s.innerHTML = svg(tipo) + "<span></span>";
    s.lastChild.textContent = testo;
    s.className = "stato " + tipo;
    clearTimeout(timerStato);
    if (tipo === "ok") timerStato = setTimeout(function () { s.className = "stato"; }, 5000);
  }

  function disegna() {
    var filtro = $("cerca").value.trim().toLowerCase();
    var ul = $("elenco");
    ul.innerHTML = "";
    $("conteggio").textContent = libreria.length === 1 ? "1 formula" : libreria.length + " formule";
    var voci = libreria.slice().sort(function (a, b) { return a.nome.localeCompare(b.nome); })
      .filter(function (v) {
        return !filtro || v.nome.toLowerCase().indexOf(filtro) >= 0 ||
          (v.descrizione || "").toLowerCase().indexOf(filtro) >= 0;
      });
    if (!voci.length) {
      var li0 = document.createElement("li");
      li0.className = "vuoto";
      li0.textContent = libreria.length ? "Nessuna formula trovata." : "La libreria è vuota: aggiungi una formula con Nuova.";
      ul.appendChild(li0);
      return;
    }
    voci.forEach(function (v) {
      var li = document.createElement("li");
      li.className = "voce";

      var riga = document.createElement("div");
      riga.className = "riga";
      var titolo = document.createElement("div");
      titolo.className = "titolo";
      var nome = document.createElement("span");
      nome.className = "nome"; nome.textContent = v.nome;
      var pill = document.createElement("span");
      pill.className = "pillola" + (v.sintassi === "en" ? " en" : "");
      pill.textContent = v.sintassi === "en" ? "EN" : "IT";
      pill.title = v.sintassi === "en" ? "Formula in sintassi inglese (letta da un file)" : "Formula in sintassi italiana";
      titolo.appendChild(nome); titolo.appendChild(pill);
      riga.appendChild(titolo);

      var strumenti = document.createElement("div");
      strumenti.className = "strumenti";
      strumenti.appendChild(pulsante("Inserisci", "inserisci", "inserisci", function (b) { inserisciUna(v, b); }, !inExcel, "Inserisci " + v.nome + " nel file"));
      strumenti.appendChild(pulsante("", "icona", "modifica", function () { apriModulo(v); }, false, "Modifica " + v.nome));
      strumenti.appendChild(pulsante("", "icona pericolo", "elimina", function (b) { elimina(v, b); }, false, "Elimina " + v.nome));
      riga.appendChild(strumenti);
      li.appendChild(riga);

      var det = document.createElement("details");
      var sum = document.createElement("summary");
      var breve = document.createElement("span");
      breve.className = "breve";
      breve.textContent = v.descrizione || "Mostra la formula";
      sum.appendChild(breve);
      sum.insertAdjacentHTML("beforeend", svg("freccia"));
      det.appendChild(sum);
      var pre = document.createElement("pre"); pre.textContent = v.formula;
      det.appendChild(pre);
      li.appendChild(det);

      ul.appendChild(li);
    });
  }

  function pulsante(testo, classe, icona, fn, disabilitato, etichetta) {
    var b = document.createElement("button");
    b.innerHTML = (icona ? svg(icona) : "") + (testo ? "<span></span>" : "");
    if (testo) b.lastChild.textContent = testo;
    if (classe) b.className = classe;
    if (etichetta) { b.setAttribute("aria-label", etichetta); b.title = etichetta; }
    if (disabilitato) { b.disabled = true; b.title = "Disponibile solo dentro Excel"; }
    b.addEventListener("click", function () { fn(b); });
    return b;
  }

  // conferma con un secondo clic entro 4 secondi (alert e confirm non funzionano nei componenti aggiuntivi)
  function chiediConferma(chiave, bottone, testoConferma) {
    if (daConfermare[chiave]) { delete daConfermare[chiave]; return true; }
    var html = bottone.innerHTML, classe = bottone.className;
    daConfermare[chiave] = true;
    bottone.textContent = testoConferma;
    if (/icona/.test(classe)) bottone.className = "conferma";
    setTimeout(function () {
      if (daConfermare[chiave]) { delete daConfermare[chiave]; bottone.innerHTML = html; bottone.className = classe; }
    }, 4000);
    return false;
  }

  // ---------- modulo nuova / modifica ----------
  function apriModulo(v) {
    $("backup").classList.add("nascosto");
    $("modulo").classList.remove("nascosto");
    inModifica = v ? v.nome : null;
    $("moduloTitolo").textContent = v ? "Modifica " + v.nome : "Nuova formula";
    $("fNome").value = v ? v.nome : "";
    $("fDescr").value = v ? (v.descrizione || "") : "";
    $("fFormula").value = v ? v.formula : "";
    $("fNome").focus();
  }
  function chiudiModulo() { $("modulo").classList.add("nascosto"); inModifica = null; }

  function salvaModulo() {
    var nome = $("fNome").value.trim();
    var descr = $("fDescr").value.trim();
    var formula = $("fFormula").value.trim();
    if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(nome)) {
      messaggio("Nome non valido: usa lettere, numeri e trattino basso, senza spazi, iniziando con una lettera.", "errore"); return;
    }
    if (formula.charAt(0) !== "=") formula = "=" + formula;
    if (formula.toUpperCase().indexOf("LAMBDA(") < 0) {
      messaggio("La formula deve contenere LAMBDA(…).", "errore"); return;
    }
    var vecchia = inModifica ? libreria[trova(inModifica)] : null;
    var esistente = trova(nome);
    if (esistente >= 0 && (!inModifica || inModifica.toUpperCase() !== nome.toUpperCase())) {
      messaggio("Esiste già una formula chiamata " + nome + ".", "errore"); return;
    }
    var voce = { nome: nome, descrizione: descr, formula: formula, sintassi: vecchia ? (vecchia.sintassi || "it") : "it" };
    if (inModifica) libreria.splice(trova(inModifica), 1, voce); else libreria.push(voce);
    salva(); chiudiModulo(); disegna();
    messaggio("Formula " + nome + " salvata nella libreria.", "ok");
  }

  function elimina(v, b) {
    if (!chiediConferma("del:" + v.nome, b, "Confermi?")) return;
    libreria.splice(trova(v.nome), 1);
    salva(); disegna();
    messaggio("Formula " + v.nome + " eliminata dalla libreria (non dal file Excel).", "ok");
  }

  // ---------- Excel ----------
  function scrivi(context, v) {
    var nomi = context.workbook.names;
    var commento = (v.descrizione || "").slice(0, 255);
    if (v.sintassi === "en") nomi.add(v.nome, v.formula, commento);
    else nomi.addFormulaLocal(v.nome, v.formula, commento);
  }

  function inserisciUna(v, b) {
    Excel.run(function (context) {
      var esistente = context.workbook.names.getItemOrNullObject(v.nome);
      return context.sync().then(function () {
        if (!esistente.isNullObject) {
          if (!chiediConferma("ins:" + v.nome, b, "Già presente: sostituire?")) return "attesa";
          esistente.delete();
        }
        scrivi(context, v);
        return context.sync().then(function () { return "fatto"; });
      });
    }).then(function (esito) {
      if (esito === "fatto") messaggio(v.nome + " inserita nel file. Usala come =" + v.nome + "(…).", "ok");
      else messaggio(v.nome + " esiste già nel file: premi di nuovo per sostituirla.", "info");
    }).catch(function (e) { errore(e, v.nome); });
  }

  function inserisciTutte(b) {
    if (!libreria.length) { messaggio("La libreria è vuota.", "info"); return; }
    if (!chiediConferma("tutte", b, "Sostituisco quelle già presenti?")) {
      messaggio("Premi di nuovo per inserire tutte le formule: quelle con lo stesso nome nel file saranno sostituite.", "info");
      return;
    }
    var fatte = 0, errori = [];
    var catena = Promise.resolve();
    libreria.forEach(function (v) {
      catena = catena.then(function () {
        return Excel.run(function (context) {
          var es = context.workbook.names.getItemOrNullObject(v.nome);
          return context.sync().then(function () {
            if (!es.isNullObject) es.delete();
            scrivi(context, v);
            return context.sync();
          });
        }).then(function () { fatte++; })
          .catch(function () { errori.push(v.nome); });
      });
    });
    catena.then(function () {
      if (errori.length) messaggio("Inserite " + fatte + " formule. Non inserite: " + errori.join(", ") + " (controlla la sintassi).", "errore");
      else messaggio(fatte === 1 ? "Inserita 1 formula nel file." : "Inserite tutte le " + fatte + " formule nel file.", "ok");
    });
  }

  function leggiDalFile() {
    Excel.run(function (context) {
      var nomi = context.workbook.names;
      nomi.load("items/name,items/formula,items/comment");
      return context.sync().then(function () {
        var nuove = 0, gia = 0;
        nomi.items.forEach(function (n) {
          if (!n.formula || n.formula.toUpperCase().indexOf("LAMBDA(") < 0) return;
          if (trova(n.name) >= 0) { gia++; return; }
          libreria.push({ nome: n.name, descrizione: n.comment || "", formula: n.formula, sintassi: "en" });
          nuove++;
        });
        salva(); disegna();
        if (nuove) messaggio((nuove === 1 ? "Aggiunta 1 formula" : "Aggiunte " + nuove + " formule") + " dal file" + (gia ? " (" + gia + " già presenti, lasciate invariate)." : "."), "ok");
        else messaggio(gia ? "Le LAMBDA del file sono già tutte in libreria." : "Nel file non ci sono formule LAMBDA in Gestione nomi.", "info");
      });
    }).catch(function (e) { errore(e); });
  }

  function errore(e, nome) {
    var dettaglio = (e && (e.message || e.code)) ? (e.message || e.code) : String(e);
    messaggio((nome ? "Non riesco a inserire " + nome + ". " : "Operazione non riuscita. ") +
      "Controlla la formula (nomi delle funzioni e separatori). Dettaglio: " + dettaglio, "errore");
  }

  // ---------- backup ----------
  function apriBackup() {
    chiudiModulo();
    $("backup").classList.remove("nascosto");
    $("fBackup").value = JSON.stringify(libreria, null, 2);
    $("fBackup").select();
  }
  function ripristina(b) {
    var l;
    try { l = JSON.parse($("fBackup").value); } catch (e) { messaggio("Il testo incollato non è un backup valido.", "errore"); return; }
    if (!Array.isArray(l) || !l.every(function (v) { return v && v.nome && v.formula; })) {
      messaggio("Il backup non contiene formule valide.", "errore"); return;
    }
    if (!chiediConferma("ripristina", b, "Sostituisco la libreria?")) return;
    libreria = l.map(function (v) { return { nome: v.nome, descrizione: v.descrizione || "", formula: v.formula, sintassi: v.sintassi === "en" ? "en" : "it" }; });
    salva(); disegna();
    $("backup").classList.add("nascosto");
    messaggio("Libreria ripristinata: " + libreria.length + " formule.", "ok");
  }

  // ---------- avvio ----------
  function avvia() {
    carica();
    $("cerca").addEventListener("input", disegna);
    $("btnNuova").addEventListener("click", function () { apriModulo(null); });
    $("btnAnnulla").addEventListener("click", chiudiModulo);
    $("btnSalva").addEventListener("click", salvaModulo);
    $("btnTutte").addEventListener("click", function () { inserisciTutte($("btnTutte")); });
    $("btnDalFile").addEventListener("click", leggiDalFile);
    $("btnBackup").addEventListener("click", apriBackup);
    $("btnChiudiBackup").addEventListener("click", function () { $("backup").classList.add("nascosto"); });
    $("btnRipristina").addEventListener("click", function () { ripristina($("btnRipristina")); });
    if (!inExcel) {
      $("btnTutte").disabled = true; $("btnDalFile").disabled = true;
      messaggio("Anteprima fuori da Excel: puoi consultare e modificare la libreria, ma l'inserimento funziona solo dentro Excel.", "info");
    }
    disegna();
  }

  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      inExcel = !!(info && info.host === Office.HostType.Excel);
      avvia();
    });
  } else {
    document.addEventListener("DOMContentLoaded", avvia);
  }
})();
