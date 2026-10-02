# Word in Markdown

Web app locale per convertire documenti Microsoft Word `.docx` in file Markdown `.md`. L'interfaccia consente di caricare o trascinare un documento, visualizzarne il risultato, copiarlo negli appunti e scaricarlo.

## Avvio su Windows

Aprire PowerShell nella cartella dell'app ed eseguire:

```powershell
Set-Location C:\Progetto_AI\APP_MarkDown
python .\docx_to_markdown_webapp.py
```

Il server si avvia su `http://127.0.0.1:8026/` e tenta di aprire la pagina nel browser predefinito. Per arrestarlo, premere `Ctrl+C` nel terminale. La cartella di lavoro e i file temporanei sono gestiti in modo compatibile con Windows.

È possibile avviare la stessa app dalla scheda **Word in Markdown** del Launcher Dashboard. La dashboard usa la porta `8026` per rilevare il servizio già attivo o avviarlo.

## Requisiti

L'applicazione usa `http.server` e gli altri moduli della Standard Library per il server e l'interfaccia. Nell'interprete Python usato per l'avvio devono essere disponibili:

- `magika`, per identificare il tipo reale del file;
- `markitdown`, motore principale di conversione;
- `markdownify`, usato dal fallback OOXML.

Non vengono installati pacchetti automaticamente. Se l'extra DOCX di MarkItDown non è presente, l'app usa il fallback locale descritto sotto. Le altre dipendenze necessarie per Word devono essere già disponibili nell'ambiente se si desidera che MarkItDown gestisca direttamente i `.docx`.

## Utilizzo

1. Aprire `http://127.0.0.1:8026/` oppure la scheda corrispondente nel Launcher Dashboard.
2. Trascinare un documento `.docx` nell'area di caricamento o selezionarlo dal computer. La dimensione massima è 20 MB.
3. Premere **Converti in Markdown**. Il risultato appare nell'anteprima.
4. Usare **Copia** per copiare il testo o **Scarica .md** per scaricare un file con il nome del documento originale e l'estensione `.md`.

Il file temporaneo usato durante la conversione viene rimosso al termine dell'elaborazione. L'app non salva il Markdown sul server: il browser scarica il contenuto ricevuto.

## Flusso di conversione

1. L'endpoint locale riceve il documento come upload `multipart/form-data`.
2. Il file viene controllato per estensione, dimensione e struttura OOXML: deve essere un archivio ZIP DOCX con `[Content_Types].xml` e `word/document.xml`.
3. Magika deve identificarlo come documento Microsoft Word DOCX. La sola estensione non è sufficiente.
4. MarkItDown prova a convertire il file. Se segnala che manca la dipendenza opzionale per DOCX, viene attivato il fallback.
5. Il fallback legge il contenuto WordprocessingML, genera HTML essenziale e lo passa a `markdownify`.
6. La risposta JSON contiene il Markdown, il motore impiegato e il nome suggerito per il download.

Il fallback conserva paragrafi, titoli, grassetto, corsivo, barrato, collegamenti, elenchi e tabelle. Gli elementi grafici non vengono incorporati nel Markdown: se Word fornisce una descrizione alternativa, viene aggiunta una nota testuale. La formattazione avanzata e alcuni dettagli di impaginazione possono variare rispetto al documento originale.

## Limiti e controlli

- Sono accettati solo file `.docx`; i vecchi documenti `.doc` non sono supportati.
- La dimensione massima caricabile è 20 MB.
- Gli archivi con più di 5.000 elementi o con contenuto espanso oltre 100 MB vengono rifiutati.
- File vuoti, corrotti, non riconosciuti da Magika o senza contenuto convertibile generano un messaggio di errore nell'interfaccia.
- Il fallback converte gli elementi di elenco in elenchi puntati; non ricostruisce necessariamente numerazione, stili personalizzati o impaginazione Word.

## Server e API

Il server ascolta esclusivamente su `127.0.0.1`, porta `8026`; non è esposto alla rete locale.

- `GET /` serve l'interfaccia web.
- `POST /api/convert` riceve il campo multipart `file` e restituisce un JSON con `markdown`, `engine`, `download_name` e `detected_type`.

Per scegliere un'altra porta, modificare `PORT` nello script oppure importare `run_server` e invocarlo con i parametri desiderati. In caso di porta occupata, chiudere il servizio che la usa oppure configurare una porta diversa anche nella relativa voce del Launcher Dashboard.
