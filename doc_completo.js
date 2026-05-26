const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, Table, TableRow, TableCell,
  WidthType, ShadingType, LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

const W = 9026;
const mono = t => new TextRun({ text: t, font: "Courier New", size: 18 });
const bold = t => new TextRun({ text: t, bold: true, font: "Arial", size: 24 });
const norm = t => new TextRun({ text: t, font: "Arial", size: 24 });
const muted = t => new TextRun({ text: t, font: "Arial", size: 22, color: "555555", italics: true });

const h1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: "Arial", size: 34, bold: true, color: "1F3864" })], spacing: { before: 480, after: 200 } });
const h2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: "Arial", size: 28, bold: true, color: "2E5090" })], spacing: { before: 320, after: 160 } });
const h3 = t => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: t, font: "Arial", size: 24, bold: true, color: "2E5090" })], spacing: { before: 200, after: 100 } });
const p = t => new Paragraph({ children: [new TextRun({ text: t, font: "Arial", size: 24 })], spacing: { after: 120 } });
const spacer = () => new Paragraph({ children: [], spacing: { after: 200 } });
const sep = () => new Paragraph({ children: [], border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } }, spacing: { after: 240, before: 80 } });
const pb = () => new Paragraph({ children: [new PageBreak()] });

const bullet = (t, level=0) => new Paragraph({
  numbering: { reference: "bullets", level },
  children: [new TextRun({ text: t, font: "Arial", size: 24 })],
  spacing: { after: 80 }
});

const numbered = (t, level=0) => new Paragraph({
  numbering: { reference: "numbers", level },
  children: [new TextRun({ text: t, font: "Arial", size: 24 })],
  spacing: { after: 80 }
});

const infoBox = (title, lines, fill, accent) => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [W],
  rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    borders: { top: { style: BorderStyle.SINGLE, size: 2, color: accent }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 8, color: accent }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
    margins: { top: 120, bottom: 120, left: 240, right: 180 },
    children: [
      new Paragraph({ children: [new TextRun({ text: title, bold: true, font: "Arial", size: 22, color: accent })], spacing: { after: 80 } }),
      ...lines.map(l => new Paragraph({ children: [new TextRun({ text: l, font: "Arial", size: 22 })], spacing: { after: 60 } }))
    ]
  })]})],
});

const twoCol = (left, right, leftW=4400, rightW=4626) => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [leftW, rightW],
  rows: [new TableRow({ children: [
    new TableCell({ width: { size: leftW, type: WidthType.DXA }, shading: { fill: "F5F8FF", type: ShadingType.CLEAR }, borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } }, margins: { top: 100, bottom: 100, left: 180, right: 120 },
      children: left.map(l => new Paragraph({ children: [new TextRun({ text: l, font: "Arial", size: 22 })], spacing: { after: 60 } })) }),
    new TableCell({ width: { size: rightW, type: WidthType.DXA }, shading: { fill: "FAFAFA", type: ShadingType.CLEAR }, borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } }, margins: { top: 100, bottom: 100, left: 180, right: 120 },
      children: right.map(l => new Paragraph({ children: [new TextRun({ text: l, font: "Arial", size: 22 })], spacing: { after: 60 } })) }),
  ]})],
});

const codeBlock = lines => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [W],
  rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA },
    shading: { fill: "1E1E2E", type: ShadingType.CLEAR },
    borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "444466" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "444466" }, left: { style: BorderStyle.SINGLE, size: 1, color: "444466" }, right: { style: BorderStyle.SINGLE, size: 1, color: "444466" } },
    margins: { top: 120, bottom: 120, left: 220, right: 160 },
    children: lines.map(l => new Paragraph({ children: [new TextRun({ text: l, font: "Courier New", size: 18, color: "CDD6F4" })], spacing: { after: 0, before: 0 } }))
  })]})],
});

/* ── DOCUMENTO RELAZIONE ── */
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 34, bold: true, font: "Arial", color: "1F3864" }, paragraph: { spacing: { before: 480, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, font: "Arial", color: "2E5090" }, paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: "Arial", color: "2E5090" }, paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
    ]
  },
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }, { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ]},
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children: [

    /* ═══ COPERTINA ═══ */
    spacer(), spacer(), spacer(),
    new Paragraph({ children: [new TextRun({ text: "LABORATORIO DI INFORMATICA", font: "Arial", size: 24, color: "7A82A0", letterSpacing: 40 })], alignment: AlignmentType.CENTER, spacing: { after: 120 } }),
    new Paragraph({ children: [new TextRun({ text: "Relazione Progettuale", font: "Arial", size: 48, bold: true, color: "1F3864" })], alignment: AlignmentType.CENTER, spacing: { after: 120 } }),
    new Paragraph({ children: [new TextRun({ text: "Sistema di Gestione per un Centro di Assistenza Tecnica", font: "Arial", size: 28, color: "2E5090", italics: true })], alignment: AlignmentType.CENTER, spacing: { after: 600 } }),
    sep(),
    spacer(),
    new Paragraph({ children: [bold("Gruppo: "), norm("Gruppo da 3 studenti — 1\u00b0 Appello")], spacing: { after: 80 } }),
    new Paragraph({ children: [bold("Corso: "), norm("Laboratorio di Informatica")], spacing: { after: 80 } }),
    new Paragraph({ children: [bold("Docente: "), norm("Prof. Alessandro Pagano")], spacing: { after: 80 } }),
    new Paragraph({ children: [bold("Data consegna: "), norm("8 Giugno 2026")], spacing: { after: 80 } }),
    new Paragraph({ children: [bold("Linguaggio: "), norm("C (standard C99)")], spacing: { after: 80 } }),
    new Paragraph({ children: [bold("File principali: "), mono("richieste.c  tecnici.c  menu.c  html.c  report.c  statistiche.c  ordinamenti.c  file_manager.c  utils.c")], spacing: { after: 400 } }),
    sep(),
    spacer(),
    pb(),

    /* ═══ SEZ 1 — ANALISI DEL PROBLEMA ═══ */
    h1("1. Analisi del Problema"),
    p("Il progetto richiede lo sviluppo di un sistema software in linguaggio C per la gestione operativa di un centro di assistenza tecnica. Il sistema deve supportare la gestione completa del ciclo di vita delle richieste di intervento su dispositivi elettronici (smartphone, notebook, tablet, stampanti, smartwatch, console), dalla ricezione alla chiusura, con assegnazione ai tecnici disponibili."),
    p("Il contesto operativo prevede che le richieste vengano gestite da operatori che necessitano di strumenti per inserire, modificare, ricercare e filtrare i dati, produrre statistiche e report, e consultare un riepilogo visivo tramite browser."),
    spacer(),

    h2("1.1 Obiettivi del sistema"),
    bullet("Memorizzare e gestire le richieste di assistenza tecnica"),
    bullet("Gestire l\u2019anagrafica dei tecnici disponibili"),
    bullet("Associare le richieste ai tecnici con logica di suggerimento automatico"),
    bullet("Aggiornare lo stato delle lavorazioni e i costi"),
    bullet("Effettuare ricerche e filtri per stato, priorit\u00e0 e cliente"),
    bullet("Ordinare le richieste secondo 3 criteri distinti con Merge Sort"),
    bullet("Generare 3 report distinti (generale, operativo, tecnici)"),
    bullet("Salvare e ricaricare i dati tramite file .txt"),
    bullet("Produrre una dashboard HTML consultabile dal browser"),
    spacer(),

    h2("1.2 Entit\u00e0 principali"),
    p("Il sistema gestisce due entit\u00e0 principali interconnesse:"),
    spacer(),
    twoCol(
      ["RICHIESTA DI ASSISTENZA", "", "Codice identificativo (int)", "Nome cliente (stringa)", "Tipologia dispositivo (stringa)", "Descrizione problema (stringa)", "Priorit\u00e0: 1=Bassa, 2=Media, 3=Alta", "Stato: Aperta / In lavorazione /", "       Completata / Annullata", "Costo stimato (float)", "Costo finale (float)", "Data di apertura (struct Data)", "Codice tecnico assegnato (int)"],
      ["TECNICO", "", "Codice identificativo (int)", "Nome (stringa)", "Area di specializzazione (stringa)", "Numero massimo richieste gestibili (int)", "Numero richieste attualmente assegnate (int)", "", "", "Il tecnico viene collegato alla richiesta", "tramite il campo codiceTecnico.", "Il valore 0 indica richiesta non assegnata."]
    ),
    spacer(), spacer(),

    h2("1.3 Flusso dei dati"),
    infoBox("Flusso operativo del sistema", [
      "File richieste.txt + file tecnici.txt",
      "            \u2193",
      "Caricamento in memoria (array statici)",
      "            \u2193",
      "Gestione richieste e tecnici (inserimento, modifica, ricerca, filtri)",
      "            \u2193",
      "Assegnazioni tecnico \u2194 richiesta (manuale e automatica)",
      "            \u2193",
      "Elaborazioni: ordinamenti, statistiche",
      "            \u2193",
      "Salvataggio dati aggiornati su file",
      "            \u2193",
      "Generazione report .txt e dashboard .html"
    ], "EBF3FB", "2E5090"),
    spacer(),
    pb(),

    /* ═══ SEZ 2 — STRUTTURE DATI ═══ */
    h1("2. Strutture Dati"),

    h2("2.1 Struct Data"),
    codeBlock([
      "typedef struct {",
      "    int giorno;",
      "    int mese;",
      "    int anno;",
      "} Data;"
    ]),
    spacer(),

    h2("2.2 Struct Richiesta"),
    codeBlock([
      "typedef struct {",
      "    int   codice;",
      "    char  cliente[MAX_NOME];         /* 50 caratteri */",
      "    char  dispositivo[MAX_DISPOSITIVO]; /* 30 caratteri */",
      "    char  problema[MAX_PROBLEMA];    /* 200 caratteri */",
      "    int   priorita;                  /* 1=Bassa 2=Media 3=Alta */",
      "    char  stato[MAX_STATO];          /* 30 caratteri */",
      "    float costoStimato;",
      "    float costoFinale;",
      "    Data  dataApertura;",
      "    int   codiceTecnico;             /* 0 = non assegnata */",
      "} Richiesta;"
    ]),
    spacer(),

    h2("2.3 Struct Tecnico"),
    codeBlock([
      "typedef struct {",
      "    int  codice;",
      "    char nome[MAX_NOME];                   /* 50 caratteri */",
      "    char specializzazione[MAX_SPECIALIZZAZIONE]; /* 50 caratteri */",
      "    int  maxRichieste;",
      "    int  richiesteAssegnate;",
      "} Tecnico;"
    ]),
    spacer(),

    h2("2.4 Costanti globali (costanti.h)"),
    codeBlock([
      "#define MAX_RICHIESTE     1000",
      "#define MAX_TECNICI       100",
      "#define MAX_NOME          50",
      "#define MAX_DISPOSITIVO   30",
      "#define MAX_PROBLEMA      200",
      "#define MAX_STATO         30",
      "#define MAX_SPECIALIZZAZIONE 50",
      "",
      "#define PRIORITA_BASSA    1",
      "#define PRIORITA_MEDIA    2",
      "#define PRIORITA_ALTA     3",
      "",
      "#define STATO_APERTA       \"Aperta\"",
      "#define STATO_LAVORAZIONE  \"In lavorazione\"",
      "#define STATO_COMPLETATA   \"Completata\"",
      "#define STATO_ANNULLATA    \"Annullata\"",
      "",
      "#define MIN_CODICE        1",
      "#define MAX_CODICE        9999",
      "",
      "#define FILE_RICHIESTE    \"data/richieste.txt\"",
      "#define FILE_TECNICI      \"data/tecnici.txt\"",
      "#define FILE_REPORT_RICHIESTE  \"output/report_richieste.txt\"",
      "#define FILE_REPORT_OPERATIVO  \"output/report_operativo.txt\"",
      "#define FILE_REPORT_TECNICI    \"output/report_tecnici.txt\"",
      "#define DASHBOARD_FILE         \"dashboard.html\""
    ]),
    spacer(),

    h2("2.5 Scelta degli array statici"),
    p("Per la gestione della memoria si \u00e8 optato per array statici di dimensione fissa (MAX_RICHIESTE = 1000, MAX_TECNICI = 100) anzich\u00e9 allocazione dinamica con malloc/realloc. Questa scelta \u00e8 motivata da:"),
    bullet("Semplicit\u00e0 implementativa: nessun rischio di memory leak o dangling pointer"),
    bullet("Stabilit\u00e0: il comportamento \u00e8 prevedibile e pi\u00f9 facile da testare"),
    bullet("Adeguatezza al contesto: le dimensioni massime sono compatibili con l\u2019uso reale"),
    bullet("Programmazione difensiva facilitata: i controlli di overflow sono immediati (if nRichieste >= MAX_RICHIESTE)"),
    spacer(),
    pb(),

    /* ═══ SEZ 3 — MODULARITA ═══ */
    h1("3. Architettura e Modularit\u00e0"),

    h2("3.1 Struttura delle cartelle"),
    codeBlock([
      "CentroAssistenza/",
      "\u251c\u2500\u2500 main.c",
      "\u251c\u2500\u2500 include/",
      "\u2502   \u251c\u2500\u2500 costanti.h       (costanti globali, stati, priorita)",
      "\u2502   \u251c\u2500\u2500 richieste.h      (struct Richiesta + prototipi)",
      "\u2502   \u251c\u2500\u2500 tecnici.h        (struct Tecnico + prototipi)",
      "\u2502   \u251c\u2500\u2500 file_manager.h   (caricamento e salvataggio)",
      "\u2502   \u251c\u2500\u2500 menu.h           (navigazione testuale)",
      "\u2502   \u251c\u2500\u2500 ordinamenti.h    (merge sort e funzioni confronto)",
      "\u2502   \u251c\u2500\u2500 statistiche.h    (calcoli statistici)",
      "\u2502   \u251c\u2500\u2500 report.h         (generazione report .txt)",
      "\u2502   \u251c\u2500\u2500 html.h           (generazione dashboard .html)",
      "\u2502   \u2514\u2500\u2500 utils.h          (funzioni comuni: leggiIntero, pulisciBuffer)",
      "\u251c\u2500\u2500 src/",
      "\u2502   \u251c\u2500\u2500 richieste.c      tecnici.c    file_manager.c",
      "\u2502   \u251c\u2500\u2500 menu.c           ordinamenti.c statistiche.c",
      "\u2502   \u251c\u2500\u2500 report.c         html.c       utils.c",
      "\u251c\u2500\u2500 data/",
      "\u2502   \u251c\u2500\u2500 richieste.txt",
      "\u2502   \u2514\u2500\u2500 tecnici.txt",
      "\u2514\u2500\u2500 output/",
      "    \u251c\u2500\u2500 report_richieste.txt",
      "    \u251c\u2500\u2500 report_operativo.txt",
      "    \u251c\u2500\u2500 report_tecnici.txt",
      "    \u2514\u2500\u2500 dashboard.html"
    ]),
    spacer(),

    h2("3.2 Responsabilit\u00e0 dei moduli"),
    p("Ogni modulo ha una responsabilit\u00e0 unica e ben definita, seguendo il principio di separazione delle preoccupazioni (separation of concerns):"),
    spacer(),
    infoBox("richieste.c", ["Inserimento, modifica, cancellazione, visualizzazione richieste.", "Ricerca per codice. Aggiornamento stato e costo finale. Assegnazione tecnico."], "F3F9EE", "2D6A2D"),
    spacer(),
    infoBox("tecnici.c", ["Inserimento, modifica, cancellazione, visualizzazione tecnici.", "Ricerca per codice. Suggerimento automatico del tecnico pi\u00f9 adatto."], "F3F9EE", "2D6A2D"),
    spacer(),
    infoBox("file_manager.c", ["Lettura da richieste.txt e tecnici.txt all\u2019avvio. Scrittura su file dopo modifiche.", "Formato CSV con separatore ';'. Validazione apertura file con controllo NULL."], "FFF8EE", "B07020"),
    spacer(),
    infoBox("ordinamenti.c", ["Implementazione Merge Sort con funzione di confronto passata come puntatore a funzione.", "Tre criteri: costo stimato (obbligatorio), priorit\u00e0, data di apertura."], "EBF3FB", "2E5090"),
    spacer(),
    infoBox("statistiche.c", ["Costo medio, costo massimo, conteggi per stato e priorit\u00e0.", "Statistiche per tipo dispositivo e per livello di priorit\u00e0."], "EBF3FB", "2E5090"),
    spacer(),
    infoBox("report.c", ["Generazione di 3 report su file .txt: generale, operativo, tecnici.", "Richiama funzioni di statistiche.c e ordinamenti.c."], "F5F0FF", "5B30A0"),
    spacer(),
    infoBox("html.c", ["Generazione di dashboard.html con CSS inline, tema scuro, 9 sezioni.", "Salvataggio automatico dei dati prima della generazione. Apertura cross-platform."], "F5F0FF", "5B30A0"),
    spacer(),
    infoBox("menu.c", ["Interfaccia testuale gerarchica progettata con AI. Loop do-while + switch-case.", "13 funzioni: 1 menu principale (9 opzioni) + 8 sottomenu + 2 sotto-sottomenu static."], "F0F0F0", "444444"),
    spacer(),
    infoBox("utils.c", ["leggiIntero(min, max): legge e valida input intero nel range specificato.", "pulisciBuffer(): svuota stdin. statoValido(stato): valida stringa stato."], "F0F0F0", "444444"),
    spacer(),
    pb(),

    /* ═══ SEZ 4 — FUNZIONALITA ═══ */
    h1("4. Funzionalit\u00e0 Implementate"),
    p("Il programma implementa tutte le 24 funzionalit\u00e0 richieste dalla traccia per gruppi da 3."),
    spacer(),

    h2("4.1 Gestione richieste"),
    numbered("Caricamento richieste da file richieste.txt all\u2019avvio"),
    numbered("Inserimento nuova richiesta con validazione di tutti i campi"),
    numbered("Modifica: stato, descrizione problema, costo stimato"),
    numbered("Cancellazione richiesta con aggiornamento contatori tecnico"),
    numbered("Visualizzazione completa di tutte le richieste"),
    numbered("Ricerca per codice con restituzione indice (-1 se non trovata)"),
    numbered("Aggiornamento costo finale delle richieste completate"),
    spacer(),

    h2("4.2 Gestione tecnici e assegnazioni"),
    numbered("Caricamento tecnici da file tecnici.txt all\u2019avvio"),
    numbered("Inserimento, modifica, cancellazione, visualizzazione tecnici"),
    numbered("Associazione manuale richiesta a tecnico con controllo maxRichieste"),
    numbered("Modifica assegnazione esistente"),
    numbered("Visualizzazione richieste non assegnate (codiceTecnico == 0)"),
    numbered("Suggerimento automatico: tecnico con stessa specializzazione, minore carico, non al massimo"),
    spacer(),

    h2("4.3 Ricerca, filtri e ordinamenti"),
    numbered("Filtra per stato (validazione con ciclo do-while)"),
    numbered("Filtra per priorit\u00e0 (range dinamico da costanti)"),
    numbered("Filtra per cliente (ricerca parziale)"),
    numbered("Ordina per costo stimato (criterio obbligatorio)"),
    numbered("Ordina per priorit\u00e0 (criterio aggiuntivo 1)"),
    numbered("Ordina per data di apertura (criterio aggiuntivo 2)"),
    spacer(),

    h2("4.4 Statistiche"),
    numbered("Costo medio stimato su tutte le richieste"),
    numbered("Costo massimo"),
    numbered("Conteggio richieste per stato"),
    numbered("Conteggio richieste per priorit\u00e0"),
    numbered("Statistiche per tipo dispositivo (PC, Smartphone, Elettrodomestici)"),
    numbered("Statistiche per livello di priorit\u00e0"),
    numbered("Statistiche tecnici: richieste assegnate e completate per tecnico"),
    spacer(),

    h2("4.5 Salvataggio"),
    numbered("Salvataggio richieste su file (manuale da menu o automatico prima della dashboard)"),
    numbered("Salvataggio tecnici su file"),
    numbered("Salvataggio selettivo o totale"),
    spacer(),
    pb(),

    /* ═══ SEZ 5 — ORDINAMENTI ═══ */
    h1("5. Algoritmo di Ordinamento: Merge Sort"),

    h2("5.1 Scelta e motivazione"),
    p("Per l\u2019ordinamento delle richieste \u00e8 stato scelto il Merge Sort. La traccia richiede obbligatoriamente uno tra Quick Sort, Merge Sort e Shell Sort. La scelta del Merge Sort \u00e8 motivata dai seguenti fattori:"),
    spacer(),
    twoCol(
      ["VANTAGGI DEL MERGE SORT", "", "\u2022 Complessit\u00e0 O(n log n) garantita nel caso peggiore", "\u2022 Comportamento stabile: mantiene l\u2019ordine relativo", "  degli elementi uguali", "\u2022 Adatto agli array statici di dimensione nota", "\u2022 Indipendente dalla distribuzione iniziale dei dati", "\u2022 Facile da estendere a diversi criteri di confronto"],
      ["CONFRONTO CON LE ALTERNATIVE", "", "Quick Sort: O(n log n) in media ma O(n\u00b2) nel caso", "peggiore su array gi\u00e0 ordinati.", "", "Shell Sort: pi\u00f9 semplice ma complessit\u00e0 dipendente", "dalla sequenza di gap scelta.", "", "Il Merge Sort \u00e8 la scelta pi\u00f9 prevedibile e robusta."]
    ),
    spacer(),

    h2("5.2 Implementazione"),
    p("Il Merge Sort \u00e8 implementato in ordinamenti.c con una funzione generica che accetta un puntatore a funzione di confronto, permettendo di riutilizzare lo stesso algoritmo per tutti e tre i criteri:"),
    codeBlock([
      "/* Firma della funzione principale */",
      "void mergeSort(Richiesta arr[], int sin, int des,",
      "               int (*confronta)(Richiesta a, Richiesta b));",
      "",
      "/* Funzioni di confronto per i 3 criteri */",
      "int confrontaCosto(Richiesta a, Richiesta b);     /* criterio obbligatorio */",
      "int confrontaPriorita(Richiesta a, Richiesta b);  /* criterio aggiuntivo 1 */",
      "int confrontaData(Richiesta a, Richiesta b);      /* criterio aggiuntivo 2 */",
      "",
      "/* Chiamata nel menu: */",
      "mergeSort(richieste, 0, nRichieste - 1, confrontaCosto);"
    ]),
    spacer(),
    pb(),

    /* ═══ SEZ 6 — REPORT ═══ */
    h1("6. Report Generati"),
    p("Il sistema genera 3 report distinti su file .txt nella cartella output/, come richiesto dalla traccia per gruppi da 3."),
    spacer(),

    h2("6.1 Report generale richieste (report_richieste.txt)"),
    bullet("Numero totale di richieste"),
    bullet("Numero di richieste per stato (Aperta, In lavorazione, Completata, Annullata)"),
    bullet("Numero di richieste per priorit\u00e0 (Bassa, Media, Alta)"),
    bullet("Costo medio stimato"),
    bullet("Costo medio finale delle richieste completate"),
    bullet("Elenco delle richieste ordinate per costo stimato"),
    bullet("Statistiche per tipo dispositivo e per priorit\u00e0"),
    spacer(),

    h2("6.2 Report operativo (report_operativo.txt)"),
    bullet("Richieste ad alta priorit\u00e0 ancora aperte"),
    bullet("Richieste non ancora assegnate"),
    bullet("Richieste in lavorazione"),
    bullet("Richieste completate"),
    bullet("Richieste annullate"),
    bullet("Riepilogo per tipologia di dispositivo (conteggi)"),
    bullet("Statistiche per dispositivo (costo medio, costo massimo)"),
    spacer(),

    h2("6.3 Report tecnici e assegnazioni (report_tecnici.txt)"),
    bullet("Elenco completo dei tecnici"),
    bullet("Numero di richieste assegnate a ciascun tecnico"),
    bullet("Dettaglio delle richieste assegnate per tecnico"),
    bullet("Tecnico con pi\u00f9 richieste assegnate"),
    bullet("Tecnico con pi\u00f9 richieste completate"),
    bullet("Richieste non assegnate"),
    bullet("Suggerimenti automatici di assegnazione"),
    spacer(),
    pb(),

    /* ═══ SEZ 7 — DASHBOARD ═══ */
    h1("7. Dashboard HTML"),

    h2("7.1 Struttura e contenuto"),
    p("Il modulo html.c genera un file dashboard.html autocontenuto (CSS inline, nessuna dipendenza esterna) consultabile da qualsiasi browser. Il file viene generato dal menu principale alla voce 9 (Dashboard HTML) con due opzioni: generazione manuale o generazione con apertura automatica nel browser."),
    p("La dashboard \u00e8 organizzata in 9 sezioni:"),
    numbered("Riepilogo generale delle richieste (KPI: totale, aperte, in lavorazione, completate, annullate)"),
    numbered("Riepilogo dei tecnici (tabella: codice, nome, specializzazione)"),
    numbered("Richieste urgenti (priorit\u00e0 alta, con badge stato colorati)"),
    numbered("Richieste non assegnate (KPI + tabella con priorit\u00e0 colorata)"),
    numbered("Stato delle assegnazioni (tabella per tecnico con barra di carico)"),
    numbered("Statistiche principali (KPI: priorit\u00e0, costo medio, costo max, % assegnazione)"),
    numbered("Report completo richieste (tabella con tutti i campi)"),
    numbered("Report operativo (richieste attive: Aperta + In lavorazione)"),
    numbered("Report tecnici (per tecnico: assegnate, completate, attive, costo totale)"),
    spacer(),

    h2("7.2 Salvataggio automatico prima della generazione"),
    p("La funzione generaDashboardHTML() salva automaticamente i dati prima di generare l\u2019HTML, garantendo coerenza tra dati in memoria e dashboard visualizzata:"),
    codeBlock([
      "int generaDashboardHTML(...) {",
      "    /* 1. Salvataggio dati aggiornati */",
      "    salvaTecnici(tecnici, nTecnici);",
      "    salvaRichieste(richieste, nRichieste);",
      "",
      "    /* 2. Salvataggio report su file .txt */",
      "    salvaReportRichieste(richieste, nRichieste);",
      "    salvaReportOperativo(richieste, nRichieste);",
      "    salvaReportTecnici(richieste, nRichieste, tecnici, nTecnici);",
      "",
      "    /* 3. Generazione HTML */",
      "    FILE *fp = fopen(DASHBOARD_FILE, \"w\");",
      "    /* ... scrittura sezioni ... */",
      "}"
    ]),
    spacer(),

    h2("7.3 Apertura cross-platform"),
    codeBlock([
      "void apriDashboardHTML(void) {",
      "#if defined(_WIN32) || defined(_WIN64)",
      "    system(\"start dashboard.html\");",
      "#elif defined(__APPLE__)",
      "    system(\"open dashboard.html\");",
      "#else",
      "    system(\"xdg-open dashboard.html 2>/dev/null\");",
      "#endif",
      "}"
    ]),
    spacer(),
    pb(),

    /* ═══ SEZ 8 — INTERFACCIA ═══ */
    h1("8. Interfaccia Utente"),

    h2("8.1 Struttura del menu (progettata con AI)"),
    p("L\u2019interfaccia testuale \u00e8 stata progettata con il supporto dell\u2019AI (Claude, Anthropic). Il sistema di navigazione \u00e8 organizzato in menu gerarchici do-while + switch-case, implementati in menu.c con 13 funzioni. La documentazione completa della sessione AI \u00e8 riportata nel documento separato Documentazione_Menu_AI."),
    codeBlock([
      "MENU PRINCIPALE (9 opzioni + 0 Esci)",
      "  1. Gestione richieste    \u2192 6 opzioni (inserisci, modifica, cancella, visualizza,",
      "                                          cerca per codice, aggiorna costo)",
      "  2. Gestione tecnici      \u2192 4 opzioni CRUD",
      "  3. Gestione assegnazioni \u2192 4 opzioni (assegna, modifica, automatico, non assegnate)",
      "  4. Ricerca e filtri      \u2192 3 opzioni (stato, priorita, cliente)",
      "  5. Ordinamenti           \u2192 3 opzioni (costo, priorita, data)",
      "  6. Statistiche           \u2192 4 opzioni + 2 sotto-sottomenu [static]",
      "                              dispositivi (3 opzioni) e priorita (3 opzioni)",
      "  7. Report                \u2192 3 opzioni (generale, operativo, tecnici)",
      "  8. Salvataggio dati      \u2192 3 opzioni (richieste, tecnici, tutto)",
      "  9. Dashboard HTML        \u2192 2 opzioni (genera, genera e apri)",
      "  0. Esci"
    ]),
    spacer(),

    h2("8.2 Scelte stilistiche dell\u2019interfaccia"),
    bullet("Separatori grafici uniformi tramite macro #define SEP per coerenza visiva"),
    bullet("Messaggi di conferma uniformi: tutti i case 0 stampano 'Ritorno al menu principale.'"),
    bullet("Messaggi di errore con range specifico: 'Inserire un numero tra 0 e N'"),
    bullet("Prompt informativi con range dinamico: printf(\"  Priorita (%d-%d): \", PRIORITA_BASSA, PRIORITA_ALTA)"),
    bullet("Validazione stato con ciclo do-while e lista degli stati validi in caso di errore"),
    spacer(),
    pb(),

    /* ═══ SEZ 9 — DIFENSIVA ═══ */
    h1("9. Programmazione Difensiva"),
    p("La traccia richiede esplicitamente la gestione di errori e input non validi. Il sistema implementa i seguenti controlli:"),
    spacer(),

    h2("9.1 Validazione input utente"),
    codeBlock([
      "/* leggiIntero: ciclo finche' l'input non e' nel range valido */",
      "int leggiIntero(int min, int max) {",
      "    int valore;",
      "    do {",
      "        scanf(\"%d\", &valore);",
      "        pulisciBuffer();",
      "    } while (valore < min || valore > max);",
      "    return valore;",
      "}",
      "",
      "/* Validazione stato con lista errori */",
      "if (!statoValido(stato)) {",
      "    printf(\"  Stati validi: %s, %s, %s, %s.\\n\",",
      "           STATO_APERTA, STATO_LAVORAZIONE,",
      "           STATO_COMPLETATA, STATO_ANNULLATA);",
      "}"
    ]),
    spacer(),

    h2("9.2 Controllo apertura file"),
    codeBlock([
      "FILE *fp = fopen(FILE_RICHIESTE, \"r\");",
      "if (fp == NULL) {",
      "    printf(\"Errore: impossibile aprire %s\\n\", FILE_RICHIESTE);",
      "    return 0;",
      "}"
    ]),
    spacer(),

    h2("9.3 Controllo array pieni"),
    codeBlock([
      "if (*nRichieste >= MAX_RICHIESTE) {",
      "    printf(\"  Errore: numero massimo di richieste raggiunto.\\n\");",
      "    return;",
      "}"
    ]),
    spacer(),

    h2("9.4 Gestione ricerche fallite"),
    codeBlock([
      "int pos = cercaRichiestaPerCodice(richieste, *nRichieste, codice);",
      "if (pos == -1) {",
      "    printf(\"\\n  Codice richiesta non trovato.\\n\");",
      "    return;",
      "}"
    ]),
    spacer(),

    h2("9.5 Controllo assegnazione tecnico"),
    codeBlock([
      "/* Verifica che il tecnico non abbia raggiunto il massimo */",
      "if (tecnici[pos].richiesteAssegnate >= tecnici[pos].maxRichieste) {",
      "    printf(\"  Tecnico al massimo della capacita'.\\n\");",
      "    return;",
      "}"
    ]),
    spacer(),
    pb(),

    /* ═══ SEZ 10 — GESTIONE FILE ═══ */
    h1("10. Gestione dei File"),

    h2("10.1 Formato file di input"),
    p("I file usano il separatore ';' per separare i campi. Questa scelta evita ambiguit\u00e0 con spazi presenti nei nomi e nelle descrizioni."),
    spacer(),
    infoBox("richieste.txt — formato", [
      "codice;cliente;dispositivo;problema;priorita;stato;costoStimato;costoFinale;gg/mm/aaaa;codiceTecnico",
      "",
      "Esempio:",
      "1;Mario Rossi;Smartphone;Schermo rotto;3;Aperta;120.50;0.00;12/05/2026;0",
      "2;Luca Verdi;Notebook;Non si accende;2;In lavorazione;250.00;0.00;10/05/2026;101"
    ], "FFF8EE", "B07020"),
    spacer(),
    infoBox("tecnici.txt — formato", [
      "codice;nome;specializzazione;maxRichieste;richiesteAssegnate",
      "",
      "Esempio:",
      "101;Giuseppe Esposito;Notebook;5;1",
      "102;Anna Bianchi;Smartphone;4;0"
    ], "FFF8EE", "B07020"),
    spacer(),

    h2("10.2 Lettura con fscanf"),
    p("La lettura del formato %d;%49[^;]; (specificatore %[^;]) permette di leggere stringhe fino al carattere ';' escludendo il delimitatore stesso, gestendo correttamente nomi e descrizioni con spazi."),
    spacer(),

    h2("10.3 Flusso di persistenza"),
    bullet("All\u2019avvio: caricaTecnici() poi caricaRichieste() (i tecnici prima perch\u00e9 le richieste li referenziano)"),
    bullet("Durante l\u2019esecuzione: salvataggio manuale dal menu 8 o automatico prima della dashboard"),
    bullet("Alla chiusura: l\u2019utente pu\u00f2 scegliere di salvare o uscire senza salvare"),
    spacer(),
    pb(),

    /* ═══ SEZ 11 — USO AI ═══ */
    h1("11. Documentazione Uso dell\u2019AI"),

    h2("11.1 Strumento utilizzato"),
    p("Strumento: Claude (Anthropic) \u2014 claude.ai. L\u2019AI \u00e8 stata utilizzata per progettare e generare il codice relativo all\u2019interfaccia testuale (menu.h, menu.c) e alla dashboard HTML (html.h, html.c), come richiesto dalla traccia."),
    spacer(),

    h2("11.2 Cosa ha progettato l\u2019AI"),
    bullet("Schermata iniziale (schermataBenvenuto) con elenco funzionalit\u00e0"),
    bullet("Struttura gerarchica del menu principale e di tutti i sottomenu"),
    bullet("Organizzazione do-while + switch-case per ogni livello di navigazione"),
    bullet("Messaggi di errore uniformi con range specifico per ogni sottomenu"),
    bullet("Messaggi di conferma uniformi al case 0 di ogni sottomenu"),
    bullet("Sotto-sottomenu per statistiche dispositivi e priorit\u00e0 (funzioni static)"),
    bullet("File html.h con dichiarazioni e html.c con CSS inline e 9 sezioni"),
    bullet("Apertura cross-platform della dashboard (Windows/macOS/Linux)"),
    spacer(),

    h2("11.3 Approccio iterativo"),
    p("La generazione non \u00e8 avvenuta in un unico prompt ma in 19 fasi iterative: analisi delle specifiche, generazione dell\u2019header, correzioni puntuali (firme errate, range mancanti, messaggi inconsistenti, bug del buffer stdin), aggiunta delle sezioni report alla dashboard, integrazione del salvataggio automatico. La documentazione completa di tutti i prompt e le risposte \u00e8 riportata nel documento separato Documentazione_Menu_AI.docx."),
    spacer(),
    pb(),

    /* ═══ SEZ 12 — CASI DI TEST ═══ */
    h1("12. Casi di Test"),

    h2("12.1 Test funzionali principali"),
    spacer(),
    new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: [700, 3000, 2800, 2526],
      rows: [
        new TableRow({ children: [
          new TableCell({ width: { size: 700, type: WidthType.DXA }, shading: { fill: "1F3864", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "#", bold: true, font: "Arial", size: 20, color: "FFFFFF" })] })] }),
          new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: "1F3864", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "Test", bold: true, font: "Arial", size: 20, color: "FFFFFF" })] })] }),
          new TableCell({ width: { size: 2800, type: WidthType.DXA }, shading: { fill: "1F3864", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "Input", bold: true, font: "Arial", size: 20, color: "FFFFFF" })] })] }),
          new TableCell({ width: { size: 2526, type: WidthType.DXA }, shading: { fill: "1F3864", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "Risultato atteso", bold: true, font: "Arial", size: 20, color: "FFFFFF" })] })] }),
        ]}),
        ...[
          ["1", "Caricamento file richieste.txt", "File valido con 5 righe", "Array richieste con 5 elementi"],
          ["2", "Caricamento file mancante", "FILE_RICHIESTE non esiste", "Messaggio errore, nRichieste = 0"],
          ["3", "Inserimento richiesta", "Tutti i campi validi", "Richiesta aggiunta, nRichieste++"],
          ["4", "Ricerca codice esistente", "Codice presente nell\u2019array", "Indice >= 0 restituito"],
          ["5", "Ricerca codice inesistente", "Codice assente", "Indice -1 restituito"],
          ["6", "Filtro per stato invalido", "Stringa non valida", "Ciclo do-while ripete, lista stati"],
          ["7", "Ordinamento per costo", "Array con costi misti", "Array ordinato crescente"],
          ["8", "Ordinamento per priorit\u00e0", "Array con priorit\u00e0 miste", "Array ordinato 1\u21922\u21923"],
          ["9", "Assegna tecnico al massimo", "richiesteAssegnate == maxRichieste", "Messaggio errore, no assegnazione"],
          ["10", "Generazione dashboard.html", "Dati validi in memoria", "File HTML creato correttamente"],
          ["11", "Input fuori range (menu)", "Numero > 9 nel menu principale", "Messaggio errore, range 0-9"],
          ["12", "Array richieste pieno", "nRichieste == MAX_RICHIESTE", "Messaggio errore, inserimento bloccato"],
          ["13", "Report generale", "Dati validi", "File report_richieste.txt creato"],
          ["14", "Salvataggio e ricaricamento", "Modifica + salva + ricarica", "Dati persistenti dopo il riavvio"],
        ].map(([n, test, input, atteso]) => new TableRow({ children: [
          new TableCell({ width: { size: 700, type: WidthType.DXA }, shading: { fill: n % 2 === 0 ? "F5F8FF" : "FAFAFA", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 80 }, borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } }, children: [new Paragraph({ children: [new TextRun({ text: n, font: "Arial", size: 20 })] })] }),
          new TableCell({ width: { size: 3000, type: WidthType.DXA }, shading: { fill: n % 2 === 0 ? "F5F8FF" : "FAFAFA", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 80 }, borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } }, children: [new Paragraph({ children: [new TextRun({ text: test, font: "Arial", size: 20 })] })] }),
          new TableCell({ width: { size: 2800, type: WidthType.DXA }, shading: { fill: n % 2 === 0 ? "F5F8FF" : "FAFAFA", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 80 }, borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } }, children: [new Paragraph({ children: [new TextRun({ text: input, font: "Courier New", size: 18 })] })] }),
          new TableCell({ width: { size: 2526, type: WidthType.DXA }, shading: { fill: n % 2 === 0 ? "F5F8FF" : "FAFAFA", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 120, right: 80 }, borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" } }, children: [new Paragraph({ children: [new TextRun({ text: atteso, font: "Arial", size: 20 })] })] }),
        ]}))
      ]
    }),
    spacer(),
    pb(),

    /* ═══ SEZ 13 — COMPILAZIONE ═══ */
    h1("13. Istruzioni di Compilazione ed Esecuzione"),

    h2("13.1 Compilazione"),
    codeBlock([
      "# Linux / macOS",
      "gcc -Wall -Wextra -std=c99 \\",
      "    src/richieste.c src/tecnici.c src/file_manager.c \\",
      "    src/menu.c src/ordinamenti.c src/statistiche.c \\",
      "    src/report.c src/html.c src/utils.c \\",
      "    main.c -o centro_assistenza",
      "",
      "# Windows (MinGW)",
      "gcc -Wall -std=c99 src\\*.c main.c -o centro_assistenza.exe",
      "",
      "# Con Makefile",
      "make"
    ]),
    spacer(),

    h2("13.2 Esecuzione"),
    codeBlock([
      "# Linux / macOS",
      "./centro_assistenza",
      "",
      "# Windows",
      "centro_assistenza.exe",
      "",
      "# I file data/richieste.txt e data/tecnici.txt devono",
      "# esistere prima dell'avvio (possono essere vuoti)."
    ]),
    spacer(),
    pb(),

    /* ═══ SEZ 14 — CONCLUSIONI ═══ */
    h1("14. Conclusioni"),
    p("Il progetto realizza un sistema completo di gestione per un centro di assistenza tecnica che soddisfa tutti i requisiti della traccia per gruppi da 3:"),
    spacer(),
    bullet("Strutture dati appropriate (struct Richiesta, struct Tecnico, struct Data)"),
    bullet("Programmazione modulare con 10 moduli separati (.h e .c)"),
    bullet("Lettura e scrittura da file .txt con formato CSV a separatore ';'"),
    bullet("24 funzionalit\u00e0 implementate (inserimento, modifica, cancellazione, ricerca, filtri, assegnazioni, suggerimento automatico)"),
    bullet("Merge Sort con 3 criteri di ordinamento (costo, priorit\u00e0, data)"),
    bullet("3 report distinti generati su file .txt"),
    bullet("Dashboard HTML autocontenuta con 9 sezioni, CSS inline, apertura cross-platform"),
    bullet("Interfaccia testuale gerarchica progettata con AI (documentata in Documentazione_Menu_AI.docx)"),
    bullet("Programmazione difensiva: validazione input, controllo file, controllo array pieni, gestione ricerche fallite"),
    spacer(),
    p("L\u2019approccio iterativo adottato, sia nello sviluppo che nella sessione con l\u2019AI, ha permesso di costruire il sistema in modo incrementale e verificabile, riducendo il rischio di errori architetturali difficili da correggere."),

  ]}]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('Relazione_Progettuale.docx', buf);
}).catch(e => { console.error(e); process.exit(1); });
