const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, Table, TableRow, TableCell,
  WidthType, ShadingType, LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

/* ── helpers ── */
const W = 9026; // content width A4 1" margins DXA
const mono = t => new TextRun({ text: t, font: "Courier New", size: 18 });
const bold = t => new TextRun({ text: t, bold: true, font: "Arial", size: 24 });
const norm = t => new TextRun({ text: t, font: "Arial", size: 24 });
const ital = t => new TextRun({ text: t, italics: true, font: "Arial", size: 22, color: "555555" });

const h1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: "Arial", size: 32, bold: true, color: "1F3864" })], spacing: { before: 400, after: 200 } });
const h2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: "Arial", size: 28, bold: true, color: "2E5090" })], spacing: { before: 300, after: 160 } });
const h3 = t => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: t, font: "Arial", size: 26, bold: true, color: "2E5090" })], spacing: { before: 200, after: 120 } });
const p = t => new Paragraph({ children: [new TextRun({ text: t, font: "Arial", size: 24 })], spacing: { after: 120 } });
const spacer = () => new Paragraph({ children: [], spacing: { after: 160 } });
const sep = () => new Paragraph({ children: [], border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } }, spacing: { after: 240, before: 80 } });
const pb = () => new Paragraph({ children: [new PageBreak()] });

const bullet = t => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  children: [new TextRun({ text: t, font: "Arial", size: 24 })],
  spacing: { after: 80 }
});

/* dialog box — utente (verde) o AI (blu) */
const dialogBox = (role, lines) => {
  const isAI = role === "AI";
  const fill = isAI ? "EBF3FB" : "F3F9EE";
  const lc = isAI ? "2E5090" : "2D6A2D";
  const bc = isAI ? "2E5090" : "2D6A2D";
  const allLines = [{ label: role, isLabel: true }, ...lines.map(l => ({ label: l, isLabel: false }))];
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      shading: { fill, type: ShadingType.CLEAR },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 2, color: bc },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
        left: { style: BorderStyle.SINGLE, size: 8, color: bc },
        right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }
      },
      margins: { top: 100, bottom: 100, left: 220, right: 160 },
      children: allLines.map(({ label, isLabel }) => new Paragraph({
        children: [new TextRun({
          text: label,
          font: isLabel ? "Arial" : "Arial",
          size: isLabel ? 20 : 22,
          bold: isLabel,
          color: isLabel ? lc : "222222"
        })],
        spacing: { after: isLabel ? 60 : 40 }
      }))
    })]})],
  });
};

/* code block dark */
const codeBlock = lines => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [W],
  rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA },
    shading: { fill: "1E1E2E", type: ShadingType.CLEAR },
    borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "444466" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "444466" }, left: { style: BorderStyle.SINGLE, size: 1, color: "444466" }, right: { style: BorderStyle.SINGLE, size: 1, color: "444466" } },
    margins: { top: 120, bottom: 120, left: 220, right: 160 },
    children: lines.map(l => new Paragraph({
      children: [new TextRun({ text: l, font: "Courier New", size: 18, color: "CDD6F4" })],
      spacing: { after: 0, before: 0 }
    }))
  })]})],
});

/* ── DOCUMENTO ── */
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, font: "Arial", color: "1F3864" }, paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, font: "Arial", color: "2E5090" }, paragraph: { spacing: { before: 300, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: "Arial", color: "2E5090" }, paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } } , children: [

    /* ═══ COPERTINA ═══ */
    spacer(), spacer(),
    new Paragraph({ children: [new TextRun({ text: "CENTRO ASSISTENZA TECNICA", font: "Arial", size: 44, bold: true, color: "1F3864" })], alignment: AlignmentType.CENTER, spacing: { after: 160 } }),
    new Paragraph({ children: [new TextRun({ text: "Progetto di Laboratorio di Informatica in C", font: "Arial", size: 28, color: "2E5090" })], alignment: AlignmentType.CENTER, spacing: { after: 400 } }),
    sep(),
    new Paragraph({ children: [new TextRun({ text: "Documentazione della sessione di generazione del menu e della dashboard HTML", font: "Arial", size: 24, italics: true, color: "444444" })], alignment: AlignmentType.CENTER, spacing: { after: 120 } }),
    new Paragraph({ children: [new TextRun({ text: "mediante Intelligenza Artificiale", font: "Arial", size: 24, italics: true, color: "444444" })], alignment: AlignmentType.CENTER, spacing: { after: 500 } }),
    new Paragraph({ children: [bold("Strumento AI: "), norm("Claude (Anthropic) \u2014 claude.ai")], spacing: { after: 80 } }),
    new Paragraph({ children: [bold("File generati: "), mono("menu.h  |  menu.c  |  html.h  |  html.c  |  main.c")], spacing: { after: 80 } }),
    new Paragraph({ children: [bold("Fase: "), norm("Progettazione e sviluppo iterativo del sistema di menu e della dashboard HTML")], spacing: { after: 400 } }),
    sep(), spacer(),

    /* ═══ SEZ 1 ═══ */
    h1("1. Introduzione e contesto del progetto"),
    p("Nell\u2019ambito del progetto universitario per il corso di Laboratorio di informatica, \u00e8 stata sviluppata un\u2019applicazione in linguaggio C per la gestione di un Centro Assistenza Tecnica. Il software gestisce richieste di assistenza, tecnici, assegnazioni, filtri, ordinamenti, statistiche, report e salvataggio dei dati su file."),
    p("Una fase del processo di sviluppo ha previsto l\u2019utilizzo di un sistema di Intelligenza Artificiale generativa (Claude) per supportare la progettazione e la generazione del codice relativo all\u2019interfaccia testuale di navigazione e alla dashboard HTML di consultazione dei dati. Il documento riporta la trascrizione completa e commentata dell\u2019intera sessione di interazione, includendo prompt, risposte, correzioni e decisioni progettuali."),
    h2("1.1 Obiettivo della sessione"),
    p("L\u2019obiettivo principale era ristrutturare il file main.c gi\u00e0 esistente, che conteneva un unico men\u00f9 con 32 opzioni numerate, trasformandolo in un sistema gerarchico a sottomenu separati organizzati per area funzionale. I file target erano: menu.h (dichiarazioni), menu.c (implementazione completa) e un main.c semplificato."),
    p("In una seconda fase, \u00e8 stato progettato e implementato il modulo di generazione della dashboard HTML (html.h e html.c), integrato come nono sottomenu nel menu principale. La sessione ha avuto un andamento iterativo con correzioni progressive."),
    h2("1.2 Materiale fornito come contesto"),
    bullet("Un documento di specifica contenente l\u2019analisi dell\u2019interfaccia testuale: struttura dei menu, messaggi di conferma/errore, stile grafico, organizzazione degli switch-case e gerarchia dei sottomen\u00f9."),
    bullet("Il file main.c originale del progetto, contenente la logica di navigazione piatta da ristrutturare e i riferimenti a tutte le funzioni gi\u00e0 implementate negli altri moduli (richieste.h, tecnici.h, statistiche.h, ordinamenti.h, report.h, file_manager.h, utils.h)."),
    spacer(),

    /* ═══ SEZ 2 ═══ */
    h1("2. Sessione di interazione con l\u2019AI"),
    p("Di seguito \u00e8 riportata la trascrizione completa della sessione. Ogni scambio \u00e8 identificato dal ruolo (UTENTE o AI). La sessione si \u00e8 articolata in diciannove fasi: le prime dieci riguardano il sistema di menu, le successive nove la dashboard HTML e le correzioni iterative."),
    spacer(),

    /* --- FASI 1-10 menu --- */
    h2("2.1 Fase 1 \u2014 Analisi delle specifiche di interfaccia"),
    dialogBox("UTENTE", ["Sto lavorando a un progetto per l\u2019esame di Laboratorio di Informatica: si tratta di un gestionale in C per un Centro Assistenza Tecnica. Ti fornisco un documento con le specifiche complete dell\u2019interfaccia testuale che il programma dovr\u00e0 avere: struttura dei menu, organizzazione dei sottomenu, messaggi di sistema e stile grafico. Prima di generare qualsiasi codice, ti chiedo di analizzare attentamente il documento e di confermare la tua comprensione della struttura richiesta."]),
    spacer(),
    dialogBox("AI", ["Ho letto e analizzato attentamente il documento. Architettura generale: programma C da terminale con una schermata iniziale seguita da un menu principale che si ramifica in 8 sottomenu funzionali, ciascuno con un proprio ciclo do-while e switch-case.", "I sottomenu identificati: Gestione Richieste (6 voci), Gestione Tecnici (4 voci), Gestione Assegnazioni (4 voci), Ricerca e Filtri (3 voci), Ordinamenti (3 voci), Statistiche con sotto-sottomenu annidati, Report (3 voci), Salvataggio Dati (3 voci).", "Pattern implementativo: do-while per i loop, switch-case per la navigazione, funzioni dedicate per ogni sottomenu. Confermo la comprensione."]),
    spacer(), sep(),

    h2("2.2 Fase 2 \u2014 Fornitura del main.c originale e analisi del codice"),
    dialogBox("UTENTE", ["Bene, la comprensione \u00e8 corretta. Ora ti fornisco il file main.c attuale del progetto: contiene un unico ciclo do-while con 32 case numerati in sequenza, tutto inline senza alcuna separazione in sottomenu. Il tuo compito sar\u00e0 generare menu.h e menu.c che ristrutturano questa navigazione nella gerarchia descritta. Prima di procedere, analizza il main.c e dimmi se hai individuato ambiguit\u00e0 che richiedono una decisione progettuale."]),
    spacer(),
    dialogBox("AI", ["Ho analizzato il main.c. Ho identificato una criticit\u00e0 che richiede una tua decisione prima di generare il codice.", "Il problema riguarda due case con logica di input inline: case 10 (aggiorna costo finale) che legge il codice richiesta con leggiIntero() e case 15 (filtra per stato) con un ciclo do-while di validazione con fgets() e statoValido().", "Le opzioni architetturali sono: A) Tenerla tutta in menu.c; B) Spostarla nei moduli funzionali esistenti; C) Incapsularla in funzioni helper private static in menu.c."]),
    spacer(), sep(),

    h2("2.3 Fase 3 \u2014 Decisione architetturale sulla gestione dell\u2019input"),
    dialogBox("UTENTE", ["Scegli la soluzione A: tutta la logica di lettura dell\u2019input rimane in menu.c. Preferisco che il modulo menu sia indipendente nella gestione dell\u2019interazione con l\u2019utente, senza dover modificare i moduli funzionali esistenti che sono gi\u00e0 stati testati."]),
    spacer(),
    dialogBox("AI", ["Capito, \u00e8 una scelta architetturalmente solida: menu.c sar\u00e0 l\u2019unico responsabile della presentazione, della lettura dell\u2019input e della validazione, mentre i moduli funzionali restano invariati.", "Ti anticipo anche che i due sotto-sottomenu delle statistiche (menuStatDisp e menuStatPrio) li implementer\u00f2 come funzioni static in menu.c, non dichiarate nell\u2019header, poich\u00e9 sono richiamate solo internamente da menuStatistiche()."]),
    spacer(), sep(),

    h2("2.4 Fase 4 \u2014 Prima generazione di menu.h"),
    dialogBox("UTENTE", ["Perfetto. Inizia dalla generazione di menu.h: voglio vedere prima l\u2019header cos\u00ec posso verificare che le chiamate alle funzioni siano corrette e coerenti con i moduli esistenti prima di procedere con l\u2019implementazione."]),
    spacer(),
    dialogBox("AI", ["Ecco menu.h. Ho strutturato il file con: guardia di inclusione, include degli header necessari (costanti.h, richieste.h, tecnici.h), commento ad albero della gerarchia di navigazione, 10 dichiarazioni pubbliche. Le due funzioni static (menuStatDisp, menuStatPrio) non compaiono perch\u00e9 private al modulo.", "Per le firme ho seguito la distinzione: parametri modificati passati come int*, parametri in sola lettura come int."]),
    spacer(),
    codeBlock([
      "/* GENERATO INTERAMENTE DALL'IA */",
      "#ifndef MENU_H",
      "#define MENU_H",
      "#include \"costanti.h\"",
      "#include \"richieste.h\"",
      "#include \"tecnici.h\"",
      "/* Struttura: schermataBenvenuto() / menuPrincipale()",
      " *   +-- menuRichieste()  +-- menuTecnici()  +-- menuAssegnazioni()",
      " *   +-- menuFiltri()     +-- menuOrdinamenti()  +-- menuStatistiche()",
      " *   +-- menuReport()     +-- menuSalvataggio()  +-- menuDashboard() */",
      "void schermataBenvenuto(void);",
      "void menuPrincipale(Richiesta richieste[], int *nRichieste,",
      "                    Tecnico   tecnici[],  int *nTecnici);",
      "void menuRichieste(Richiesta richieste[], int *nRichieste,",
      "                   Tecnico   tecnici[],  int *nTecnici);",
      "/* ... altre dichiarazioni ... */",
      "#endif /* MENU_H */"
    ]),
    spacer(), sep(),

    h2("2.5 Fase 5 \u2014 Revisione dell\u2019header e correzione di una firma"),
    dialogBox("UTENTE", ["Noto un problema nelle chiamate di menuRichieste() e menuTecnici(). Nel caso di menuRichieste, il parametro nTecnici dovrebbe essere passato come int e non come int*, perch\u00e9 la funzione non modifica il numero di tecnici: li usa solo in sola lettura. Correggilo."]),
    spacer(),
    dialogBox("AI", ["Correzione immediata. La firma errata era:", "  void menuRichieste(Richiesta richieste[], int *nRichieste, Tecnico tecnici[], int *nTecnici);", "La firma corretta \u00e8:", "  void menuRichieste(Richiesta richieste[], int *nRichieste, Tecnico tecnici[], int nTecnici);", "Stessa logica simmetrica per menuTecnici(), dove nRichieste \u00e8 passato come int."]),
    spacer(),
    codeBlock([
      "/* Firma corretta dopo la revisione */",
      "void menuRichieste(Richiesta richieste[], int *nRichieste,",
      "                   Tecnico   tecnici[],  int  nTecnici);   /* int, non int* */",
      "",
      "void menuTecnici(Tecnico   tecnici[],  int *nTecnici,",
      "                 Richiesta richieste[], int  nRichieste);  /* int, non int* */"
    ]),
    spacer(), sep(),

    h2("2.6 Fase 6 \u2014 Generazione di menu.c"),
    dialogBox("UTENTE", ["Adesso \u00e8 corretto. Procedi con la generazione completa di menu.c, assicurandoti di usare esattamente le chiamate aggiornate e di includere tutti i sottomen\u00f9 discussi, compresi i due sotto-sottomen\u00f9 per le statistiche."]),
    spacer(),
    dialogBox("AI", ["Ecco menu.c completo. Il file implementa 12 funzioni: schermataBenvenuto(), menuPrincipale() con 8 opzioni, menuRichieste() con 6 opzioni, menuTecnici() con 4 opzioni CRUD, menuAssegnazioni(), menuFiltri() con validazione stato, menuOrdinamenti(), menuStatistiche(), [static] menuStatDisp(), [static] menuStatPrio(), menuReport(), menuSalvataggio().", "Ho usato due macro locali (#define SEP e #define LINE) per i separatori grafici, garantendo uniformit\u00e0 visiva senza ripetizioni."]),
    spacer(),
    codeBlock([
      "/* GENERATO INTERAMENTE DALL'IA */",
      "#define SEP  \"========================================================\\n\"",
      "#define LINE \"--------------------------------------------------------\\n\"",
      "",
      "void menuPrincipale(Richiesta richieste[], int *nRichieste,",
      "                    Tecnico   tecnici[],  int *nTecnici)",
      "{",
      "    int scelta = -1;",
      "    do {",
      "        printf(SEP);",
      "        printf(\"                   MENU PRINCIPALE\\n\");",
      "        /* ... 8 opzioni + 0 Esci ... */",
      "        scelta = leggiIntero(0, 8);",
      "        switch (scelta) {",
      "        case 1: menuRichieste(...); break;",
      "        /* ... */",
      "        case 0: printf(\"\\n  Uscita in corso...\\n\"); break;",
      "        default: printf(\"Scelta non valida. Inserire un numero tra 0 e 8.\\n\");",
      "        }",
      "    } while (scelta != 0);",
      "}"
    ]),
    spacer(), sep(),

    h2("2.7 Fase 7 \u2014 Revisione del menuFiltri e range dinamico"),
    dialogBox("UTENTE", ["Ho esaminato il menuFiltri(). Il ciclo di validazione per lo stato funziona correttamente, ma nel case 2 (filtra per priorit\u00e0) non stai stampando all\u2019utente il range di valori accettabili prima di chiamare leggiIntero(). Correggi questo aspetto."]),
    spacer(),
    dialogBox("AI", ["Correzione applicata. Il prompt nel case 2 di menuFiltri \u00e8 stato aggiornato da printf(\"  Priorita: \") a printf(\"  Priorita (%d-%d): \", PRIORITA_BASSA, PRIORITA_ALTA). Il range viene mostrato dinamicamente usando le costanti definite in costanti.h, anzich\u00e9 valori hardcoded."]),
    spacer(),
    codeBlock([
      "case 2:",
      "{",
      "    int priorita = 0;",
      "    /* Range mostrato dinamicamente con le costanti di costanti.h */",
      "    printf(\"  Priorita (%d-%d): \", PRIORITA_BASSA, PRIORITA_ALTA);",
      "    priorita = leggiIntero(PRIORITA_BASSA, PRIORITA_ALTA);",
      "    filtraPerPriorita(richieste, nRichieste, priorita);",
      "    break;",
      "}"
    ]),
    spacer(), sep(),

    h2("2.8 Fase 8 \u2014 Uniformit\u00e0 dei messaggi di ritorno al menu"),
    dialogBox("UTENTE", ["Nel documento di specifica era indicato che ogni sottomen\u00f9 dovrebbe mostrare un messaggio quando l\u2019utente sceglie 0 per tornare al menu principale. Guardando il codice generato, alcuni case 0 stampano il messaggio e altri no. Vorrei che tutti mostrassero 'Ritorno al menu principale.'"]),
    spacer(),
    dialogBox("AI", ["Correzione applicata a tutti i sottomenu. I case 0 che mancavano del messaggio erano: menuTecnici, menuAssegnazioni, menuOrdinamenti, menuReport.", "Il messaggio uniformato in tutti i case 0 \u00e8:", "  printf(\"\\n  Ritorno al menu principale.\\n\");", "Per i sotto-sottomenu (menuStatDisp e menuStatPrio) il messaggio \u00e8 'Ritorno alle statistiche.' per essere contestualmente pi\u00f9 preciso."]),
    spacer(), sep(),

    h2("2.9 Fase 9 \u2014 Uniformit\u00e0 dei messaggi di errore nei default"),
    dialogBox("UTENTE", ["Altro errore: i messaggi nei case default degli switch non sono uniformi. In alcuni sottomenu il messaggio specifica il range, in altri dice solo 'Scelta non valida.' Vorrei che tutti i default riportassero il range specifico di quel sottomenu."]),
    spacer(),
    dialogBox("AI", ["Correzione applicata a tutti i case default. Ogni sottomenu ora mostra il range specifico:", "  menuPrincipale  -> 'Scelta non valida. Inserire un numero tra 0 e 8.'", "  menuRichieste   -> 'Scelta non valida. Inserire un numero tra 0 e 6.'", "  menuTecnici     -> 'Scelta non valida. Inserire un numero tra 0 e 4.'", "  menuFiltri      -> 'Scelta non valida. Inserire un numero tra 0 e 3.'", "  menuOrdinamenti -> 'Scelta non valida. Inserire un numero tra 0 e 3.'", "  menuStatistiche -> 'Scelta non valida. Inserire un numero tra 0 e 4.'", "  menuReport      -> 'Scelta non valida. Inserire un numero tra 0 e 3.'", "  menuSalvataggio -> 'Scelta non valida. Inserire un numero tra 0 e 3.'"]),
    spacer(), sep(),

    h2("2.10 Fase 10 \u2014 Generazione del main.c e problema del buffer stdin"),
    dialogBox("UTENTE", ["Il codice di menu.h e menu.c \u00e8 ora corretto. Genera il main.c aggiornato che utilizzi il nuovo sistema di menu, scrivendo solo l\u2019inizializzazione degli array, caricamento dati da file e avvio del menu."]),
    spacer(),
    dialogBox("AI", ["Ecco il main.c aggiornato. Il file si riduce a quattro operazioni: dichiarazione degli array, caricamento dati persistenti (caricaTecnici, caricaRichieste), chiamata a schermataBenvenuto(), avvio del loop principale con menuPrincipale(). Include solo gli header strettamente necessari."]),
    spacer(),
    dialogBox("UTENTE", ["Perfetto. Un\u2019ultima cosa: nella schermataBenvenuto() la funzione usa getchar() per attendere l\u2019INVIO, ma se il programma viene avviato subito dopo un input da terminale potrebbe esserci un '\\n' residuo nel buffer che fa saltare l\u2019attesa. Come si gestisce?"]),
    spacer(),
    dialogBox("AI", ["Osservazione corretta. La soluzione \u00e8 svuotare il buffer prima di chiamare getchar(). Si pu\u00f2 usare la funzione pulisciBuffer() gi\u00e0 presente in utils.h, per coerenza con il codebase."]),
    spacer(),
    dialogBox("UTENTE", ["Usa pulisciBuffer() per coerenza. Aggiorna schermataBenvenuto() di conseguenza."]),
    spacer(),
    dialogBox("AI", ["Aggiornamento applicato. La prima istruzione di schermataBenvenuto() \u00e8 ora pulisciBuffer(), seguita dalla stampa della schermata e dal getchar() finale."]),
    spacer(),
    codeBlock([
      "void schermataBenvenuto(void)",
      "{",
      "    pulisciBuffer();  /* svuota buffer residui da stdin */",
      "    printf(\"\\n\");",
      "    printf(SEP);",
      "    printf(\"             CENTRO ASSISTENZA TECNICA\\n\");",
      "    /* ... resto invariato ... */",
      "    getchar();",
      "}"
    ]),
    spacer(), sep(),

    /* ═══ FASI 11-19: DASHBOARD HTML ═══ */
    pb(),
    h2("2.11 Fase 11 \u2014 Richiesta del modulo dashboard HTML"),
    dialogBox("UTENTE", ["Il progetto prevede anche la possibilit\u00e0 di esportare i dati in una dashboard HTML consultabile dal browser. Voglio aggiungere due nuovi file: html.h e html.c.", "La dashboard deve includere almeno:", "  - riepilogo generale delle richieste;", "  - riepilogo dei tecnici;", "  - richieste urgenti (priorita massima);", "  - richieste non assegnate;", "  - stato delle assegnazioni per tecnico;", "  - statistiche principali;", "  - report completo di tutte le richieste.", "La visualizzazione non deve essere interattiva, ma deve essere leggibile e organizzata. Inizia da html.h."]),
    spacer(),
    dialogBox("AI", ["Ho analizzato i requisiti. Propongo due funzioni pubbliche in html.h:", "  1. generaDashboardHTML(richieste[], nRichieste, tecnici[], nTecnici): funzione principale che genera il file dashboard.html. Restituisce 0 in caso di successo, -1 in caso di errore.", "  2. apriDashboardHTML(void): tenta l\u2019apertura automatica del file nel browser tramite system(), con fallback su messaggio per apertura manuale. Usa direttive #if defined per gestire Windows, macOS e Linux.", "Tutte le funzioni di scrittura interna sono static in html.c e non compaiono nell\u2019header. Confermo la struttura."]),
    spacer(),
    codeBlock([
      "/* GENERATO INTERAMENTE DALL'IA */",
      "#ifndef HTML_H",
      "#define HTML_H",
      "#include \"costanti.h\"",
      "#include \"richieste.h\"",
      "#include \"tecnici.h\"",
      "",
      "#define DASHBOARD_FILE \"dashboard.html\"",
      "",
      "/* Genera dashboard.html con tutti i dati del sistema.",
      " * Ritorna 0 se ok, -1 in caso di errore. */",
      "int  generaDashboardHTML(Richiesta richieste[], int nRichieste,",
      "                         Tecnico   tecnici[],  int nTecnici);",
      "",
      "/* Apre dashboard.html nel browser predefinito (best-effort). */",
      "void apriDashboardHTML(void);",
      "",
      "#endif /* HTML_H */"
    ]),
    spacer(), sep(),

    h2("2.12 Fase 12 \u2014 Generazione di html.c"),
    dialogBox("UTENTE", ["Le firme sono corrette. Procedi con la generazione di html.c completo."]),
    spacer(),
    dialogBox("AI", ["Ecco html.c. Le scelte implementative principali sono:", "  - Macro W(fp, ...): alias di fprintf per ridurre la verbosit\u00e0 delle circa 200 chiamate di scrittura.", "  - CSS inline nel <head>: tema scuro con variabili CSS, tabelle responsive, badge colorati per stato e priorit\u00e0, barre di carico per le assegnazioni.", "  - Timestamp: la data di generazione \u00e8 inclusa nell\u2019header HTML tramite time() e strftime().", "  - 9 funzioni static private, una per ciascuna sezione della dashboard.", "  - Compatibilit\u00e0: nessuna dipendenza esterna; il file \u00e8 autocontenuto e compila con gcc standard."]),
    spacer(),
    codeBlock([
      "#include <stdio.h>",
      "#include <stdlib.h>",
      "#include <string.h>",
      "#include <time.h>",
      "#include \"../include/html.h\"",
      "#include \"../include/file_manager.h\"",
      "",
      "#define W(fp, ...)  fprintf((fp), __VA_ARGS__)",
      "",
      "static void scriviHead(FILE *fp);",
      "static void scriviHeader(FILE *fp);",
      "static void scriviSezioneRiepilogoRichieste(FILE*, Richiesta[], int);",
      "static void scriviSezioneRiepilogoTecnici(FILE*, Tecnico[], int);",
      "static void scriviSezioneUrgenti(FILE*, Richiesta[], int);",
      "static void scriviSezioneNonAssegnate(FILE*, Richiesta[], int);",
      "static void scriviSezioneAssegnazioni(FILE*, Richiesta[], int, Tecnico[], int);",
      "static void scriviSezioneStatistiche(FILE*, Richiesta[], int, Tecnico[], int);",
      "static void scriviSezioneReport(FILE*, Richiesta[], int, Tecnico[], int);",
      "static void scriviSezioneReportOperativo(FILE*, Richiesta[], int);",
      "static void scriviSezioneReportTecnici(FILE*, Richiesta[], int, Tecnico[], int);",
      "static void scriviFooter(FILE *fp);"
    ]),
    spacer(), sep(),

    h2("2.13 Fase 13 \u2014 Integrazione nel menu principale"),
    dialogBox("UTENTE", ["Ora integra il modulo html nel sistema di menu:", "  1. Aggiungi menuDashboard() come nuova funzione in menu.h e menu.c;", "  2. Aggiorna menuPrincipale() per includere l\u2019opzione 9 (Dashboard HTML);", "  3. Aggiorna il default di menuPrincipale() al nuovo range 0-9.", "Il sottomenu deve offrire due opzioni: generazione senza apertura automatica e generazione con apertura nel browser."]),
    spacer(),
    dialogBox("AI", ["Aggiornamenti applicati a menu.h e menu.c.", "menu.h: aggiunta la dichiarazione void menuDashboard(...) e aggiornato il commento ad albero con il nuovo ramo.", "menu.c - menuPrincipale(): aggiunta voce '9. Dashboard HTML', aggiunto case 9 -> menuDashboard(...), aggiornato leggiIntero(0, 9), aggiornato il default a 'tra 0 e 9', aggiunto #include html.h.", "menu.c - menuDashboard(): 2 opzioni (genera / genera e apri), verifica del valore di ritorno di generaDashboardHTML() prima di chiamare apriDashboardHTML().", "Anche schermataBenvenuto() \u00e8 stata aggiornata: aggiunta voce '- Dashboard HTML' nell\u2019elenco delle funzionalit\u00e0."]),
    spacer(),
    codeBlock([
      "/* menuPrincipale() aggiornato */",
      "printf(\"  8. Salvataggio dati\\n\");",
      "printf(\"  9. Dashboard HTML\\n\");   /* aggiunto */",
      "scelta = leggiIntero(0, 9);         /* era (0, 8) */",
      "",
      "case 9:",
      "    menuDashboard(richieste, *nRichieste, tecnici, *nTecnici);",
      "    break;",
      "",
      "default:",
      "    printf(\"\\n  Scelta non valida. Inserire un numero tra 0 e 9.\\n\");",
      "",
      "/* menuDashboard() — nuova funzione */",
      "void menuDashboard(Richiesta richieste[], int nRichieste,",
      "                   Tecnico   tecnici[],  int nTecnici)",
      "{",
      "    int scelta = -1;",
      "    do {",
      "        printf(\"  1. Genera dashboard.html\\n\");",
      "        printf(\"  2. Genera e apri nel browser\\n\");",
      "        printf(\"  0. Torna al menu principale\\n\");",
      "        scelta = leggiIntero(0, 2);",
      "        switch (scelta) {",
      "        case 1: generaDashboardHTML(...); break;",
      "        case 2: { int e = generaDashboardHTML(...); if (e==0) apriDashboardHTML(); break; }",
      "        case 0: printf(\"\\n  Ritorno al menu principale.\\n\"); break;",
      "        default: printf(\"\\n  Scelta non valida. Inserire un numero tra 0 e 2.\\n\");",
      "        }",
      "    } while (scelta != 0);",
      "}"
    ]),
    spacer(), sep(),

    h2("2.14 Fase 14 \u2014 Correzione: rimozione di pulisciBuffer() da schermataBenvenuto()"),
    dialogBox("UTENTE", ["Anche modificando il path di menu.h in main.c, non d\u00e0 errore, compila ma non parte (rimane bloccato)."]),
    spacer(),
    dialogBox("AI", ["Il problema \u00e8 quasi certamente schermataBenvenuto(): pulisciBuffer() consuma tutto l\u2019input finch\u00e9 non trova '\\n' o EOF. Se stdin \u00e8 gi\u00e0 vuoto quando arriva pulisciBuffer(), il ciclo si blocca in attesa di input e il programma sembra congelato prima ancora di mostrare qualcosa.", "Il fix: rimuovere pulisciBuffer() dall\u2019inizio di schermataBenvenuto(). A programma appena avviato stdin \u00e8 vuoto, quindi getchar() da solo \u00e8 sufficiente e corretto."]),
    spacer(),
    codeBlock([
      "/* VERSIONE CORRETTA: senza pulisciBuffer() all'avvio */",
      "void schermataBenvenuto(void)",
      "{",
      "    printf(\"\\n\");",
      "    printf(SEP);",
      "    printf(\"             CENTRO ASSISTENZA TECNICA\\n\");",
      "    printf(SEP);",
      "    /* ... resto invariato ... */",
      "    getchar();  /* attende INVIO: nessun pulisciBuffer() prima */",
      "}"
    ]),
    spacer(), sep(),

    h2("2.15 Fase 15 \u2014 Correzione: path del file dashboard.html"),
    dialogBox("UTENTE", ["Quando clicco su menu dashboard --> genera html e apri browser, esce un errore 'ERRORE: impossibile creare ../output/dashboard.html'."]),
    spacer(),
    dialogBox("AI", ["Il problema \u00e8 il path: la directory ../output/ non esiste e fopen in scrittura non crea automaticamente le directory mancanti.", "La soluzione consigliata \u00e8 usare un path semplice senza sottodirectory, modificando la macro in html.h:", "  #define DASHBOARD_FILE \"dashboard.html\"", "Il file viene creato nella directory corrente da cui si lancia il programma. Nessuna directory da creare, funziona ovunque."]),
    spacer(),
    codeBlock([
      "/* html.h — macro corretta */",
      "#define DASHBOARD_FILE \"dashboard.html\"   /* path semplice, no sottodirectory */"
    ]),
    spacer(), sep(),

    h2("2.16 Fase 16 \u2014 Rimozione della colonna cognome dalla tabella tecnici"),
    dialogBox("UTENTE", ["Rimuovi la colonna relativa al cognome dalla sezione riepilogo tecnici. La tabella dovr\u00e0 essere di tre colonne: codice, nome e specializzazione."]),
    spacer(),
    dialogBox("AI", ["Due correzioni in scriviSezioneRiepilogoTecnici(): rimosso <th>Cognome</th> dall\u2019header e rimosso tecnici[i].cognome dalla riga dati."]),
    spacer(),
    codeBlock([
      "/* Header tabella corretto */",
      "W(fp, \"<th>Codice</th><th>Nome</th><th>Specializzazione</th>\\n\");",
      "",
      "/* Riga dati corretta */",
      "W(fp, \"<tr><td>%d</td><td>%s</td><td>%s</td></tr>\\n\",",
      "      tecnici[i].codice,",
      "      tecnici[i].nome,",
      "      tecnici[i].specializzazione);"
    ]),
    spacer(), sep(),

    h2("2.17 Fase 17 \u2014 Aggiunta del report operativo e del report tecnici alla dashboard"),
    dialogBox("UTENTE", ["Aggiungi in fondo alla pagina anche la stampa del report operativo e del report tecnici."]),
    spacer(),
    dialogBox("AI", ["Aggiunte due nuove sezioni statiche in html.c e due nuove chiamate in generaDashboardHTML().", "scriviSezioneReportOperativo(): mostra le richieste in stato APERTA o IN LAVORAZIONE con un KPI del conteggio totale.", "scriviSezioneReportTecnici(): mostra per ogni tecnico: nome, specializzazione, richieste assegnate, completate, attive e costo totale stimato. Completate in verde (prio-1), attive in arancione (prio-2)."]),
    spacer(),
    codeBlock([
      "/* In generaDashboardHTML() — ordine di chiamata aggiornato */",
      "scriviSezioneReport(fp, richieste, nRichieste, tecnici, nTecnici);",
      "scriviSezioneReportOperativo(fp, richieste, nRichieste);            /* aggiunto */",
      "scriviSezioneReportTecnici(fp, richieste, nRichieste, tecnici, nTecnici); /* aggiunto */",
      "W(fp, \"</main>\\n\");"
    ]),
    spacer(), sep(),

    h2("2.18 Fase 18 \u2014 Salvataggio dati prima della generazione HTML"),
    dialogBox("UTENTE", ["Fai s\u00ec che, prima della generazione del file html, salvi i dati delle richieste e dei tecnici, richiamando le funzioni salvaTecnici() e salvaRichieste()."]),
    spacer(),
    dialogBox("AI", ["La modifica va fatta in generaDashboardHTML() in html.c, aggiungendo le due chiamate di salvataggio all\u2019inizio, prima di aprire il file HTML.", "Aggiunto anche #include \"../include/file_manager.h\" in cima a html.c.", "Il salvataggio non blocca la generazione in caso di errore: stampa un avviso ma prosegue."]),
    spacer(),
    codeBlock([
      "int generaDashboardHTML(Richiesta richieste[], int nRichieste,",
      "                        Tecnico   tecnici[],  int nTecnici)",
      "{",
      "    /* Salvataggio dati prima della generazione */",
      "    if (salvaTecnici(tecnici, nTecnici) == 1)",
      "        printf(\"\\nTecnici salvati con successo!!\");",
      "    if (salvaRichieste(richieste, nRichieste) == 1)",
      "        printf(\"\\nRichieste salvate con successo!\");",
      "    printf(\"\\n\");",
      "",
      "    /* Salvataggio dei file di report */",
      "    salvaReportRichieste(richieste, nRichieste);",
      "    salvaReportOperativo(richieste, nRichieste);",
      "    salvaReportTecnici(richieste, nRichieste, tecnici, nTecnici);",
      "",
      "    FILE *fp = fopen(DASHBOARD_FILE, \"w\");",
      "    /* ... resto invariato ... */",
      "}"
    ]),
    spacer(), sep(),

    h2("2.19 Fase 19 \u2014 Correzione del doppio tag </main> in generaDashboardHTML()"),
    dialogBox("AI", ["Durante la revisione finale del codice \u00e8 stato rilevato un tag </main> duplicato in generaDashboardHTML(), introdotto per errore durante l\u2019aggiunta delle sezioni report. Rimossa la riga duplicata."]),
    spacer(),
    codeBlock([
      "/* PRIMA (errato) */",
      "W(fp, \"</main>\\n\");",
      "W(fp, \"</main>\\n\");   /* duplicato: da rimuovere */",
      "",
      "/* DOPO (corretto) */",
      "W(fp, \"</main>\\n\");"
    ]),
    spacer(),

    /* ═══ SEZ 3 — CODICE FINALE ═══ */
    pb(),
    h1("3. Codice generato (versione finale)"),
    p("In questa sezione \u00e8 riportato il codice sorgente completo nella sua versione definitiva, dopo tutte le revisioni e correzioni documentate nella sezione 2."),

    h2("3.1 File: html.h"),
    codeBlock([
      "/* GENERATO INTERAMENTE DALL'IA */",
      "#ifndef HTML_H",
      "#define HTML_H",
      "#include \"costanti.h\"",
      "#include \"richieste.h\"",
      "#include \"tecnici.h\"",
      "",
      "/* Struttura della dashboard generata:",
      " *   generaDashboardHTML()",
      " *     +-- Riepilogo generale richieste",
      " *     +-- Riepilogo tecnici (3 colonne: codice, nome, specializzazione)",
      " *     +-- Richieste urgenti (priorita massima)",
      " *     +-- Richieste non assegnate",
      " *     +-- Stato delle assegnazioni (barra di carico per tecnico)",
      " *     +-- Statistiche principali (KPI: priorita, costi, % assegnazione)",
      " *     +-- Report completo richieste",
      " *     +-- Report operativo (richieste attive)",
      " *     +-- Report tecnici (per tecnico: assegnate/completate/attive/costo)",
      " */",
      "",
      "#define DASHBOARD_FILE \"dashboard.html\"",
      "",
      "int  generaDashboardHTML(Richiesta richieste[], int nRichieste,",
      "                         Tecnico   tecnici[],  int nTecnici);",
      "",
      "void apriDashboardHTML(void);",
      "",
      "#endif /* HTML_H */"
    ]),
    spacer(),

    h2("3.2 File: html.c (struttura e funzioni principali)"),
    p("Il file completo conta circa 500 righe. Sono riportati di seguito gli estratti pi\u00f9 significativi."),

    h3("Include, macro e prototipi privati"),
    codeBlock([
      "#include <stdio.h>",
      "#include <stdlib.h>",
      "#include <string.h>",
      "#include <time.h>",
      "#include \"../include/html.h\"",
      "#include \"../include/costanti.h\"",
      "#include \"../include/richieste.h\"",
      "#include \"../include/tecnici.h\"",
      "#include \"../include/file_manager.h\"",
      "",
      "#define W(fp, ...)  fprintf((fp), __VA_ARGS__)"
    ]),
    spacer(),

    h3("generaDashboardHTML \u2014 funzione pubblica principale"),
    codeBlock([
      "int generaDashboardHTML(Richiesta richieste[], int nRichieste,",
      "                        Tecnico   tecnici[],  int nTecnici)",
      "{",
      "    /* 1. Salvataggio dati e report su file */",
      "    if (salvaTecnici(tecnici, nTecnici) == 1)",
      "        printf(\"\\nTecnici salvati con successo!!\");",
      "    if (salvaRichieste(richieste, nRichieste) == 1)",
      "        printf(\"\\nRichieste salvate con successo!\");",
      "    printf(\"\\n\");",
      "    salvaReportRichieste(richieste, nRichieste);",
      "    salvaReportOperativo(richieste, nRichieste);",
      "    salvaReportTecnici(richieste, nRichieste, tecnici, nTecnici);",
      "",
      "    /* 2. Apertura file HTML */",
      "    FILE *fp = fopen(DASHBOARD_FILE, \"w\");",
      "    if (fp == NULL) {",
      "        printf(\"\\n  ERRORE: impossibile creare %s.\\n\", DASHBOARD_FILE);",
      "        return -1;",
      "    }",
      "",
      "    /* 3. Generazione HTML */",
      "    W(fp, \"<!DOCTYPE html>\\n<html lang=\\\"it\\\">\\n\");",
      "    scriviHead(fp);",
      "    W(fp, \"<body>\\n\");",
      "    scriviHeader(fp);",
      "    W(fp, \"<main class=\\\"container\\\">\\n\");",
      "    scriviSezioneRiepilogoRichieste(fp, richieste, nRichieste);",
      "    scriviSezioneRiepilogoTecnici(fp, tecnici, nTecnici);",
      "    scriviSezioneUrgenti(fp, richieste, nRichieste);",
      "    scriviSezioneNonAssegnate(fp, richieste, nRichieste);",
      "    scriviSezioneAssegnazioni(fp, richieste, nRichieste, tecnici, nTecnici);",
      "    scriviSezioneStatistiche(fp, richieste, nRichieste, tecnici, nTecnici);",
      "    scriviSezioneReport(fp, richieste, nRichieste, tecnici, nTecnici);",
      "    scriviSezioneReportOperativo(fp, richieste, nRichieste);",
      "    scriviSezioneReportTecnici(fp, richieste, nRichieste, tecnici, nTecnici);",
      "    W(fp, \"</main>\\n\");",
      "    scriviFooter(fp);",
      "    W(fp, \"</body>\\n</html>\\n\");",
      "    fclose(fp);",
      "",
      "    printf(\"\\n  Dashboard generata: %s\\n\", DASHBOARD_FILE);",
      "    return 0;",
      "}"
    ]),
    spacer(),

    h3("apriDashboardHTML \u2014 apertura cross-platform"),
    codeBlock([
      "void apriDashboardHTML(void)",
      "{",
      "    int esito = -1;",
      "#if defined(_WIN32) || defined(_WIN64)",
      "    esito = system(\"start dashboard.html\");",
      "#elif defined(__APPLE__)",
      "    esito = system(\"open dashboard.html\");",
      "#else",
      "    esito = system(\"xdg-open dashboard.html 2>/dev/null\");",
      "#endif",
      "    if (esito != 0)",
      "        printf(\"  Apertura automatica non riuscita.\\n\"",
      "               \"  Aprire manualmente il file: %s\\n\", DASHBOARD_FILE);",
      "}"
    ]),
    spacer(),

    h2("3.3 File: menu.h (versione finale)"),
    codeBlock([
      "/* GENERATO INTERAMENTE DALL'IA */",
      "#ifndef MENU_H",
      "#define MENU_H",
      "#include \"costanti.h\"",
      "#include \"richieste.h\"",
      "#include \"tecnici.h\"",
      "/* Struttura:",
      " *   schermataBenvenuto()        schermata iniziale",
      " *   menuPrincipale()            loop principale do-while",
      " *     +-- menuRichieste()       sottomenu richieste",
      " *     +-- menuTecnici()         sottomenu tecnici",
      " *     +-- menuAssegnazioni()    sottomenu assegnazioni",
      " *     +-- menuFiltri()          sottomenu ricerca e filtri",
      " *     +-- menuOrdinamenti()     sottomenu ordinamenti",
      " *     +-- menuStatistiche()     sottomenu statistiche",
      " *     |     +-- menuStatDisp()  sotto-sottomenu dispositivi [static]",
      " *     |     +-- menuStatPrio()  sotto-sottomenu priorita    [static]",
      " *     +-- menuReport()          sottomenu report",
      " *     +-- menuSalvataggio()     sottomenu salvataggio",
      " *     +-- menuDashboard()       sottomenu dashboard HTML",
      " */",
      "void schermataBenvenuto(void);",
      "void menuPrincipale(Richiesta richieste[], int *nRichieste,",
      "                    Tecnico   tecnici[],  int *nTecnici);",
      "void menuRichieste(Richiesta richieste[], int *nRichieste,",
      "                   Tecnico   tecnici[],  int  nTecnici);",
      "void menuTecnici(Tecnico   tecnici[],  int *nTecnici,",
      "                 Richiesta richieste[], int  nRichieste);",
      "void menuAssegnazioni(Richiesta richieste[], int nRichieste,",
      "                      Tecnico   tecnici[],  int nTecnici);",
      "void menuFiltri(Richiesta richieste[], int nRichieste);",
      "void menuOrdinamenti(Richiesta richieste[], int nRichieste);",
      "void menuStatistiche(Richiesta richieste[], int nRichieste,",
      "                     Tecnico   tecnici[],  int nTecnici);",
      "void menuReport(Richiesta richieste[], int nRichieste,",
      "                Tecnico   tecnici[],  int nTecnici);",
      "void menuSalvataggio(Richiesta richieste[], int nRichieste,",
      "                     Tecnico   tecnici[],  int nTecnici);",
      "void menuDashboard(Richiesta richieste[], int nRichieste,",
      "                   Tecnico   tecnici[],  int nTecnici);",
      "#endif /* MENU_H */"
    ]),
    spacer(),

    h2("3.4 File: menu.c \u2014 modifiche rispetto alla versione precedente"),
    p("Le modifiche alla versione originale di menu.c sono tre punti precisi."),

    h3("schermataBenvenuto() \u2014 rimosso pulisciBuffer(), aggiunta voce Dashboard HTML"),
    codeBlock([
      "void schermataBenvenuto(void)",
      "{",
      "    /* pulisciBuffer() rimosso: causa blocco se stdin e' gia' vuoto */",
      "    printf(\"\\n\");",
      "    /* ... */",
      "    printf(\"    - Salvataggio dati\\n\");",
      "    printf(\"    - Dashboard HTML\\n\");   /* aggiunto */",
      "    /* ... */",
      "    getchar();",
      "}"
    ]),
    spacer(),

    h3("menuPrincipale() \u2014 opzione 9 e range aggiornato"),
    codeBlock([
      "printf(\"  8. Salvataggio dati\\n\");",
      "printf(\"  9. Dashboard HTML\\n\");   /* aggiunto */",
      "printf(\"  0. Esci\\n\");",
      "scelta = leggiIntero(0, 9);         /* era leggiIntero(0, 8) */",
      "",
      "case 9:",
      "    menuDashboard(richieste, *nRichieste, tecnici, *nTecnici);",
      "    break;",
      "",
      "default:",
      "    printf(\"\\n  Scelta non valida. Inserire un numero tra 0 e 9.\\n\");"
    ]),
    spacer(),

    h3("menuDashboard() \u2014 nuova funzione"),
    codeBlock([
      "void menuDashboard(Richiesta richieste[], int nRichieste,",
      "                   Tecnico   tecnici[],  int nTecnici)",
      "{",
      "    int scelta = -1;",
      "    do {",
      "        printf(\"  1. Genera dashboard.html\\n\");",
      "        printf(\"  2. Genera e apri nel browser\\n\");",
      "        printf(\"  0. Torna al menu principale\\n\");",
      "        scelta = leggiIntero(0, 2);",
      "        switch (scelta) {",
      "        case 1:",
      "        {",
      "            int esito = generaDashboardHTML(richieste, nRichieste,",
      "                                            tecnici, nTecnici);",
      "            if (esito == 0)",
      "                printf(\"  Aprire manualmente: %s\\n\", DASHBOARD_FILE);",
      "            break;",
      "        }",
      "        case 2:",
      "        {",
      "            int esito = generaDashboardHTML(richieste, nRichieste,",
      "                                            tecnici, nTecnici);",
      "            if (esito == 0) apriDashboardHTML();",
      "            break;",
      "        }",
      "        case 0:",
      "            printf(\"\\n  Ritorno al menu principale.\\n\");",
      "            break;",
      "        default:",
      "            printf(\"\\n  Scelta non valida. Inserire un numero tra 0 e 2.\\n\");",
      "        }",
      "    } while (scelta != 0);",
      "}"
    ]),
    spacer(),

    /* ═══ SEZ 4 — ANALISI ═══ */
    pb(),
    h1("4. Analisi delle scelte progettuali"),
    p("Le sezioni 4.1\u20134.6 riguardano le scelte relative al sistema di menu e sono invariate rispetto alla documentazione originale. Le sezioni 4.7\u20134.12 descrivono le scelte progettuali introdotte con il modulo html."),

    h2("4.1 Separazione netta tra navigazione e logica di business"),
    p("Il principio guida dell\u2019intera sessione \u00e8 stata la separazione tra il modulo menu, responsabile della presentazione e dell\u2019interazione con l\u2019utente, e i moduli funzionali esistenti, responsabili della logica di business. Menu.c non duplica n\u00e9 riscrive nessuna funzione: si limita a invocarle passando i parametri necessari."),

    h2("4.2 Verifica iterativa anziché generazione monolitica"),
    p("La sessione ha seguito un approccio incrementale: prima l\u2019analisi delle specifiche, poi la generazione dell\u2019header (verificato e corretto prima di procedere), infine l\u2019implementazione. Questo ha ridotto il costo delle correzioni."),

    h2("4.3 Funzioni static per i sotto-sottomenu"),
    p("Le funzioni menuStatDisp() e menuStatPrio() sono dichiarate static in menu.c e non compaiono nell\u2019header. Il principio \u00e8 la minima visibilit\u00e0: non ha senso esporre nell\u2019interfaccia pubblica funzioni che sono dettagli implementativi interni."),

    h2("4.4 Macro locali per i separatori grafici"),
    p("L\u2019uso di #define SEP garantisce che il separatore grafico sia definito in un unico punto e usato uniformemente in tutte le 13 funzioni di menu.c. Modificare lo stile visivo richiede un cambiamento in una sola riga di codice."),

    h2("4.5 Prompt informativi con costanti simboliche"),
    p("A seguito della revisione della fase 7, i prompt che richiedono un input numerico con range mostrano il range usando le costanti definite in costanti.h (es. PRIORITA_BASSA, PRIORITA_ALTA) anzich\u00e9 valori hardcoded."),

    h2("4.6 Rimozione di pulisciBuffer() da schermataBenvenuto()"),
    p("L\u2019analisi della fase 14 ha evidenziato che pulisciBuffer(), chiamata all\u2019inizio di schermataBenvenuto(), bloccava il programma all\u2019avvio perch\u00e9 stdin \u00e8 gi\u00e0 vuoto e il ciclo di pulizia attendeva un carattere che non arrivava mai. La soluzione corretta \u00e8 rimuoverla: getchar() da solo \u00e8 sufficiente."),

    h2("4.7 Separazione tra generazione e presentazione in html.c"),
    p("Il modulo html.c mantiene la stessa separazione di responsabilit\u00e0 gi\u00e0 adottata nel modulo menu: le funzioni static private si occupano ciascuna di una sezione della dashboard, mentre generaDashboardHTML() coordina la sequenza di scrittura senza contenere logica di presentazione."),

    h2("4.8 CSS inline e file HTML autocontenuto"),
    p("Il file dashboard.html \u00e8 completamente autocontenuto: non richiede risorse esterne, fogli di stile separati o script JavaScript. Il CSS \u00e8 scritto nel blocco <style> generato da scriviHead(). Questo garantisce che il file sia visualizzabile correttamente da qualsiasi browser anche senza connessione di rete."),

    h2("4.9 Macro W(fp, ...) per la scrittura HTML"),
    p("L\u2019alias W(fp, ...) \u00e8 definito localmente in html.c come macro per fprintf. Riduce la verbosit\u00e0 delle circa 200 chiamate di scrittura nel file, rendendo il codice pi\u00f9 leggibile senza introdurre overhead: la macro viene espansa inline dal preprocessore C."),

    h2("4.10 Apertura cross-platform con direttive condizionali"),
    p("La funzione apriDashboardHTML() usa direttive #if defined per selezionare il comando di sistema corretto: start su Windows, open su macOS, xdg-open su Linux. In caso di errore, la funzione non termina il programma ma stampa il percorso del file per apertura manuale."),

    h2("4.11 Verifica del valore di ritorno in menuDashboard()"),
    p("Nel case 2 di menuDashboard(), apriDashboardHTML() viene chiamata solo se generaDashboardHTML() restituisce 0 (successo). Questa verifica impedisce di tentare l\u2019apertura di un file che non \u00e8 stato creato correttamente."),

    h2("4.12 Salvataggio dati integrato nella generazione della dashboard"),
    p("La funzione generaDashboardHTML() salva automaticamente i dati di tecnici e richieste su file prima di generare l\u2019HTML, richiamando salvaTecnici() e salvaRichieste(). Questo garantisce coerenza tra i dati mostrati nella dashboard e quelli persistiti su disco. Il salvataggio non blocca la generazione in caso di errore: viene stampato un avviso e il processo continua."),
    spacer(),

    /* ═══ SEZ 5 — CONCLUSIONI ═══ */
    h1("5. Conclusioni"),
    p("La sessione di interazione con l\u2019AI si \u00e8 articolata in diciannove fasi complessive, che hanno coperto l\u2019intero ciclo di sviluppo del modulo menu (fasi 1\u201310) e del modulo di generazione della dashboard HTML (fasi 11\u201319)."),
    p("Il processo ha mostrato come l\u2019utilizzo efficace di uno strumento AI generativo richieda un approccio critico e iterativo da parte del programmatore: non accettare il codice generato passivamente, ma verificarlo, identificare inconsistenze e richiedere correzioni motivate. Le fasi 5, 7, 8, 9, 14, 15, 16, 17, 18 e 19 documentano altrettanti interventi correttivi che hanno migliorato la qualit\u00e0 finale del codice rispetto alla prima generazione."),
    p("Il codice prodotto comprende cinque file definitivi: menu.h e menu.c (sistema di navigazione gerarchico a 13 funzioni), html.h e html.c (dashboard HTML autocontenuta con 9 sezioni), main.c (invariato, meno di 25 righe). La dashboard generata include tutte le sezioni richieste dalle specifiche e salva automaticamente i dati su file a ogni generazione."),

  ]}]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('Documentazione_CAT_completa.docx', buf);
}).catch(e => { console.error(e); process.exit(1); });
node /home/claude/doc_completo.js 
