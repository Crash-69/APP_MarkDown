# Ruolo e Obiettivo
Agisci come un esperto sviluppatore Full-Stack Python ed esperto di interfacce web/UX.
Il tuo obiettivo è creare una **Web App locale** per trasformare documenti da formato Word (.docx) a formato Markdown (.md).

# Contesto Tecnico e Dipendenze
- **Sistema Operativo:** Windows
- **Librerie già installate e pronte all'uso:** `markitdown`, `magika`, `markdownify`.
- **Vincolo stringente sulle dipendenze:** Non installare software o pacchetti esterni aggiuntivi. Utilizza esclusivamente le librerie già presenti sul PC, i moduli della Standard Library di Python e il framework UI ereditato dalla Fase 1.

---

# Fasi di Esecuzione

### Fase 1: Ispezione del Workspace e Analisi dello Stile
1. Esamina attentamente il file `estrattore_pdf_webapp.py` situato nella cartella `Estrae_testo_da_Pdf_richiama_Ollama`.
2. Identifica il framework UI utilizzato (es. Streamlit, Gradio, o Custom HTML/CSS con Flask/FastAPI).
3. Analizza e isola le componenti di design: la palette di colori (es. tema scuro/cyber/secure), i fogli di stile CSS personalizzati, il layout dei componenti e i font per usarli come base visiva coerente per la nuova Web App.

### Fase 2: Definizione delle Funzionalità di Conversione
Pianifica la logica di conversione sfruttando gli strumenti a disposizione:
1. Utilizza `markitdown` come motore principale per la conversione da Word a Markdown.
2. Integra `magika` per la validazione/identificazione preliminare del tipo di file inserito dall'utente.
3. Gestisci la conversione in modo robusto, prevedendo eccezioni in caso di file corrotti.

### Fase 3: Pianificazione e Architettura della Web App
Crea un piano d'azione per generare la nuova Web App in un unico file Python (o in una struttura minimale). La Web App deve:
- Ereditare fedelmente lo stile visivo, il look & feel e la libreria UI dell'estrattore PDF analizzato nella Fase 1.
- Mostrare un'interfaccia pulita, moderna e organizzata a schede (tab) o a griglia (card).
- Includere un'area di drag-and-drop o un pulsante di caricamento file per i documenti Word.
- Mostrare un'anteprima del testo Markdown convertito direttamente nell'interfaccia.
- Fornire un pulsante per scaricare direttamente il file `.md` generato.

### Fase 4: Implementazione e Generazione del Codice
- Scrivi il codice sorgente completo, modulare, pronto all'uso e interamente commentato in italiano.
- Non usare placeholder o codice troncato ("inserisci qui la logica..."). Scrivi lo script nella sua interezza.
- Assicurati che i percorsi di salvataggio dei file siano compatibili con l'ambiente Windows.

### Fase 5: Verifica ed Esecuzione
1. Crea il nuovo file di script nel workspace locale.
2. Avvialo utilizzando il terminale integrato di Antigravity.
3. Verifica il corretto rendering dell'interfaccia nel browser e testa il flusso di conversione (caricamento, elaborazione, anteprima e download) per assicurarti che non ci siano errori di runtime.
