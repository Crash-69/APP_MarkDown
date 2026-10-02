#!/usr/bin/env python3
"""Web app locale per convertire documenti DOCX in Markdown.

Interfaccia e server HTTP basati sulla standard library, in continuità con
l'estrattore PDF locale. MarkItDown è il motore principale; se nell'ambiente
manca il suo extra DOCX, viene usato un fallback OOXML locale con markdownify.

Avvio: python docx_to_markdown_webapp.py
"""

from __future__ import annotations

import email.policy
import html
import json
import os
import re
import sys
import tempfile
import threading
import time
import urllib.parse
import webbrowser
import zipfile
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.etree import ElementTree

from magika import Magika
from markitdown import MarkItDown
from markdownify import markdownify


# Si forza UTF-8 per mantenere corretti i messaggi in console su Windows.
os.environ["PYTHONIOENCODING"] = "utf-8"
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

HOST = "127.0.0.1"
PORT = 8026
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
MAX_ARCHIVE_ENTRIES = 5000
MAX_UNCOMPRESSED_BYTES = 100 * 1024 * 1024
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


# Pagina singola che riprende palette, superfici, accenti e area di upload
# dell'estrattore PDF di riferimento.
HTML_PAGE = r"""<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#f5f5f5">
<title>Word in Markdown</title>
<style>
  :root { color-scheme: light; --ink:#0f172a; --muted:#64748b; --green:#059669; --green-dark:#047857; --line:#e2e8f0; --paper:#fff; --soft:#f8fafc; }
  * { box-sizing:border-box; }
  body { margin:0; min-height:100vh; background:#f5f5f5; color:#1a1a1a; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif; }
  .wrap { width:min(1120px,100%); margin:0 auto; padding:44px 24px 28px; }
  .header { display:flex; align-items:center; justify-content:space-between; gap:24px; margin-bottom:28px; }
  .brand { display:flex; align-items:center; gap:16px; min-width:0; }
  .logo { width:58px; height:58px; flex:0 0 58px; border-radius:16px; display:grid; place-items:center; background:var(--paper); border:1px solid rgba(0,0,0,.06); box-shadow:0 4px 12px rgba(0,0,0,.04); color:var(--green-dark); font-size:13px; font-weight:800; }
  h1 { margin:0 0 5px; color:var(--ink); font-size:28px; line-height:1.2; }
  .subtitle { margin:0; color:var(--muted); font-size:14px; }
  .local-badge { display:flex; align-items:center; gap:8px; flex:0 0 auto; color:#475569; font-size:12px; font-weight:650; }
  .dot { width:8px; height:8px; border-radius:50%; background:#10b981; }
  .grid { display:grid; grid-template-columns:minmax(280px,.78fr) minmax(0,1.22fr); align-items:start; gap:20px; }
  .card { min-width:0; padding:26px; background:var(--paper); border:1px solid rgba(0,0,0,.06); border-radius:18px; box-shadow:0 2px 12px rgba(0,0,0,.03); }
  .card-head { display:flex; align-items:center; justify-content:space-between; gap:12px; margin-bottom:18px; }
  .card-title { margin:0; color:var(--ink); font-size:17px; font-weight:650; }
  .step { color:#94a3b8; font-size:11px; font-weight:700; letter-spacing:.08em; }
  .dropzone { position:relative; min-height:220px; padding:28px 18px; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; cursor:pointer; background:var(--soft); border:2px dashed #cbd5e1; border-radius:15px; transition:border-color .16s,background .16s; }
  .dropzone:hover,.dropzone.dragging,.dropzone.has-file { border-color:var(--green); background:rgba(5,150,105,.045); }
  .drop-icon { width:46px; height:46px; margin-bottom:15px; display:grid; place-items:center; border:1px solid var(--line); border-radius:13px; background:#fff; color:var(--green-dark); font-size:21px; font-weight:700; }
  .drop-title { margin:0 0 6px; color:var(--ink); font-size:15px; font-weight:650; }
  .drop-hint { margin:0; color:var(--muted); font-size:12px; line-height:1.6; }
  .browse { color:var(--green-dark); text-decoration:underline; text-underline-offset:3px; }
  .file-input { position:absolute; width:1px; height:1px; padding:0; overflow:hidden; clip:rect(0,0,0,0); white-space:nowrap; clip-path:inset(50%); }
  .file-meta { display:none; width:100%; margin-top:16px; padding-top:13px; border-top:1px solid var(--line); color:#334155; font-size:12px; overflow-wrap:anywhere; }
  .has-file .file-meta { display:block; }
    .action { width:100%; min-height:48px; margin-top:16px; display:flex; align-items:center; justify-content:center; gap:9px; border:0; border-radius:11px; background:var(--green); color:#fff; box-shadow:0 7px 17px rgba(5,150,105,.18); font:inherit; font-size:14px; font-weight:600; cursor:pointer; transition:background .15s,transform .15s; }
  .action:hover:not(:disabled) { background:var(--green-dark); }
  .action:active:not(:disabled) { transform:translateY(1px); }
  .action:disabled { background:#cbd5e1; color:#64748b; box-shadow:none; cursor:not-allowed; }
  .note { margin:13px 0 0; color:#94a3b8; font-size:11px; line-height:1.55; }
  .status { display:none; margin-top:14px; padding:12px 13px; border:1px solid #bbf7d0; border-radius:10px; background:#f0fdf4; color:#166534; font-size:12px; line-height:1.5; }
  .status.error { border-color:#fee2e2; background:#fef2f2; color:#b91c1c; }
  .status.visible { display:block; }
  .result-card { min-height:370px; }
  .result-head { margin-bottom:14px; }
  .result-actions { display:flex; align-items:center; gap:8px; }
    .icon-button { min-height:34px; padding:0 11px; display:inline-flex; align-items:center; justify-content:center; gap:7px; border:1px solid var(--line); border-radius:8px; background:#fff; color:#334155; font:inherit; font-size:12px; font-weight:600; cursor:pointer; }
  .icon-button:hover:not(:disabled) { border-color:#94a3b8; background:var(--soft); }
  .icon-button:disabled { color:#94a3b8; cursor:not-allowed; }
  .preview { min-height:290px; max-height:62vh; margin:0; padding:18px; overflow:auto; white-space:pre-wrap; overflow-wrap:anywhere; border:1px solid var(--line); border-radius:12px; background:#0f172a; color:#d1fae5; font:13px/1.7 Consolas,"Cascadia Code",monospace; tab-size:2; }
  .empty { min-height:250px; display:grid; place-items:center; text-align:center; color:#94a3b8; font-size:13px; }
  .empty-mark { margin:0 auto 10px; color:#cbd5e1; font-size:29px; }
  .footer { padding:24px 0 0; text-align:center; color:#94a3b8; font-size:11px; }
  .spin { display:inline-block; animation:spin .8s linear infinite; }
  @keyframes spin { to { transform:rotate(360deg); } }
  @media (max-width:760px) { .wrap { padding:28px 16px 20px; } .header { align-items:flex-start; } .grid { grid-template-columns:1fr; } .card { padding:21px; } .result-card { min-height:0; } .preview { min-height:260px; max-height:55vh; } }
  @media (max-width:460px) { .header { flex-direction:column; gap:13px; } h1 { font-size:24px; } .brand { gap:12px; } .logo { width:50px; height:50px; flex-basis:50px; } .result-head { align-items:flex-start; flex-direction:column; } .result-actions { width:100%; } .icon-button { flex:1; } }
  @media (prefers-reduced-motion:reduce) { *,*::before,*::after { animation-duration:.01ms !important; transition-duration:.01ms !important; } }
</style>
</head>
<body>
<main class="wrap">
  <header class="header">
    <div class="brand">
      <div class="logo" aria-hidden="true">DOCX<br>→ MD</div>
    <div><h1>File in Markdown</h1><p class="subtitle">Conversione DOCX o anteprima Markdown, in locale.</p></div>
    </div>
    <div class="local-badge"><span class="dot"></span>Elaborazione sul dispositivo</div>
  </header>
    <section class="grid" aria-label="Conversione o anteprima Markdown">
    <article class="card">
            <div class="card-head"><h2 class="card-title">File di origine</h2><span class="step">01 / CARICA</span></div>
      <label class="dropzone" id="dropzone" for="fileInput">
        <span class="drop-icon" aria-hidden="true">↑</span>
        <span class="drop-title" id="dropTitle">Trascina qui il tuo documento</span>
                <span class="drop-hint" id="dropHint">oppure <span class="browse">scegli un file</span><br>Formato .docx o .md · massimo 20 MB</span>
        <span class="file-meta" id="fileMeta"></span>
                <input class="file-input" id="fileInput" type="file" accept=".docx,.md,application/vnd.openxmlformats-officedocument.wordprocessingml.document,text/markdown,text/plain">
      </label>
            <button class="action" id="convertButton" type="button" disabled><span aria-hidden="true">↗</span> Elabora file</button>
            <p class="note">I file vengono elaborati in locale e rimossi al termine dell’elaborazione.</p>
      <div class="status" id="status" role="status" aria-live="polite"></div>
    </article>
    <article class="card result-card">
      <div class="card-head result-head">
        <div><h2 class="card-title">Anteprima Markdown</h2><span class="step">02 / RISULTATO</span></div>
        <div class="result-actions">
          <button class="icon-button" id="copyButton" type="button" disabled title="Copia il Markdown negli appunti">Copia</button>
          <button class="icon-button" id="downloadButton" type="button" disabled title="Scarica il file Markdown">↓ Scarica .md</button>
        </div>
      </div>
      <div class="empty" id="emptyState"><div><div class="empty-mark" aria-hidden="true">¶</div>Il risultato apparirà qui dopo la conversione.</div></div>
      <pre class="preview" id="preview" hidden></pre>
    </article>
  </section>
    <footer class="footer">DOCX / Markdown <span aria-hidden="true">·</span> Server locale 127.0.0.1</footer>
</main>
<script>
(() => {
  const fileInput = document.getElementById('fileInput');
  const dropzone = document.getElementById('dropzone');
  const convertButton = document.getElementById('convertButton');
  const copyButton = document.getElementById('copyButton');
  const downloadButton = document.getElementById('downloadButton');
  const preview = document.getElementById('preview');
  const emptyState = document.getElementById('emptyState');
  const statusBox = document.getElementById('status');
  const fileMeta = document.getElementById('fileMeta');
  const dropTitle = document.getElementById('dropTitle');
  const dropHint = document.getElementById('dropHint');
  let selectedFile = null;
  let markdownText = '';
  let downloadName = 'documento.md';

  function showStatus(message, isError = false) {
    statusBox.textContent = message;
    statusBox.classList.toggle('error', isError);
    statusBox.classList.add('visible');
  }

  function formatBytes(size) {
    return size < 1024 * 1024 ? `${(size / 1024).toFixed(0)} KB` : `${(size / (1024 * 1024)).toFixed(1)} MB`;
  }

  function selectFile(file) {
    if (!file) return;
        const filename = file.name.toLowerCase();
        if (!filename.endsWith('.docx') && !filename.endsWith('.md')) {
      selectedFile = null;
      fileInput.value = '';
      dropzone.classList.remove('has-file');
      convertButton.disabled = true;
            showStatus('Seleziona un file con estensione .docx o .md.', true);
      return;
    }
    if (file.size > 20 * 1024 * 1024) {
      selectedFile = null;
      fileInput.value = '';
      dropzone.classList.remove('has-file');
      convertButton.disabled = true;
      showStatus('Il file supera il limite consentito di 20 MB.', true);
      return;
    }
    selectedFile = file;
    fileMeta.textContent = `${file.name} · ${formatBytes(file.size)}`;
    dropTitle.textContent = 'File pronto per l’elaborazione';
    dropHint.textContent = 'Seleziona un altro file per sostituirlo';
    dropzone.classList.add('has-file');
    convertButton.disabled = false;
    statusBox.classList.remove('visible');
  }

  fileInput.addEventListener('change', () => selectFile(fileInput.files[0]));
  for (const eventName of ['dragenter', 'dragover']) {
    dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.add('dragging'); });
  }
  for (const eventName of ['dragleave', 'drop']) {
    dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.remove('dragging'); });
  }
  dropzone.addEventListener('drop', (event) => selectFile(event.dataTransfer.files[0]));

  convertButton.addEventListener('click', async () => {
    if (!selectedFile) return;
    const formData = new FormData();
    formData.append('file', selectedFile, selectedFile.name);
    convertButton.disabled = true;
    convertButton.innerHTML = '<span class="spin" aria-hidden="true">◌</span> Conversione in corso';
    statusBox.classList.remove('visible');
    try {
      const response = await fetch('/api/convert', { method: 'POST', body: formData });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'Conversione non riuscita.');
      markdownText = result.markdown;
      downloadName = result.download_name;
      preview.textContent = markdownText;
      preview.hidden = false;
      emptyState.hidden = true;
      copyButton.disabled = false;
      downloadButton.disabled = false;
    const engineNote = result.engine === 'MarkItDown' ? 'Convertito con MarkItDown.' : result.engine === 'Markdown' ? 'File Markdown caricato.' : 'Convertito con il fallback DOCX locale: l’extra DOCX di MarkItDown non è installato.';
      showStatus(`${engineNote} Tipo riconosciuto: ${result.detected_type}.`);
    } catch (error) {
      showStatus(error.message || 'Errore durante la conversione.', true);
    } finally {
      convertButton.disabled = !selectedFile;
    convertButton.innerHTML = '<span aria-hidden="true">↗</span> Elabora file';
    }
  });

  copyButton.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(markdownText);
      copyButton.textContent = 'Copiato';
      window.setTimeout(() => { copyButton.textContent = 'Copia'; }, 1400);
    } catch {
      showStatus('Impossibile accedere agli appunti del browser.', true);
    }
  });

  downloadButton.addEventListener('click', () => {
    const blob = new Blob([markdownText], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = downloadName;
        document.body.appendChild(anchor);
    anchor.click();
        anchor.remove();
        window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  });
})();
</script>
</body>
</html>"""


def _word_tag(name: str) -> str:
    """Restituisce il nome XML completo di un elemento WordprocessingML."""
    return f"{{{WORD_NS}}}{name}"


def validate_docx(content: bytes) -> None:
    """Controlla che il file sia uno ZIP DOCX plausibile e non un archivio enorme."""
    if not content:
        raise ValueError("Il file caricato è vuoto.")
    try:
        with zipfile.ZipFile(__import__("io").BytesIO(content)) as archive:
            entries = archive.infolist()
            if len(entries) > MAX_ARCHIVE_ENTRIES:
                raise ValueError("Il documento contiene un numero anomalo di elementi.")
            if sum(entry.file_size for entry in entries) > MAX_UNCOMPRESSED_BYTES:
                raise ValueError("Il contenuto estratto supera il limite di sicurezza.")
            names = set(archive.namelist())
            if "[Content_Types].xml" not in names or "word/document.xml" not in names:
                raise ValueError("Il file non contiene la struttura di un documento Word DOCX.")
            content_types = archive.read("[Content_Types].xml")
            if b"wordprocessingml.document.main+xml" not in content_types:
                raise ValueError("Il file non è un documento Word DOCX valido.")
            if archive.getinfo("word/document.xml").file_size > MAX_UNCOMPRESSED_BYTES:
                raise ValueError("La parte principale del documento supera il limite di sicurezza.")
    except zipfile.BadZipFile as exc:
        raise ValueError("Il documento è corrotto o non è un file DOCX valido.") from exc


def _read_relationships(archive: zipfile.ZipFile) -> dict[str, str]:
    """Legge i collegamenti esterni usati dagli hyperlink nel documento."""
    try:
        root = ElementTree.fromstring(archive.read("word/_rels/document.xml.rels"))
    except (KeyError, ElementTree.ParseError):
        return {}
    relationships: dict[str, str] = {}
    for relation in root.findall(f"{{{PACKAGE_REL_NS}}}Relationship"):
        relation_id = relation.get("Id")
        target = relation.get("Target")
        if relation_id and target and relation.get("TargetMode") == "External":
            relationships[relation_id] = target
    return relationships


def _render_run(run: ElementTree.Element, relationships: dict[str, str] | None = None) -> str:
    """Converte testo e formattazione essenziale di un run Word in HTML."""
    pieces: list[str] = []
    for node in run.iter():
        if node.tag == _word_tag("t"):
            pieces.append(html.escape(node.text or ""))
        elif node.tag in (_word_tag("tab"),):
            pieces.append(" &nbsp; ")
        elif node.tag in (_word_tag("br"), _word_tag("cr")):
            pieces.append("<br>")
        elif node.tag == _word_tag("noBreakHyphen"):
            pieces.append("&#8209;")
        elif node.tag == _word_tag("softHyphen"):
            pieces.append("&shy;")
    text = "".join(pieces)
    if not text:
        return ""

    properties = run.find(_word_tag("rPr"))
    if properties is not None:
        if properties.find(_word_tag("b")) is not None:
            text = f"<strong>{text}</strong>"
        if properties.find(_word_tag("i")) is not None:
            text = f"<em>{text}</em>"
        if properties.find(_word_tag("strike")) is not None or properties.find(_word_tag("dstrike")) is not None:
            text = f"<del>{text}</del>"
        style = properties.find(_word_tag("rStyle"))
        if style is not None and "code" in (style.get(_word_tag("val")) or "").lower():
            text = f"<code>{text}</code>"
    return text


def _paragraph_html(paragraph: ElementTree.Element, relationships: dict[str, str]) -> str:
    """Estrae contenuto, collegamenti e stile di un paragrafo Word."""
    fragments: list[str] = []
    for child in paragraph:
        if child.tag == _word_tag("r"):
            fragments.append(_render_run(child))
        elif child.tag == _word_tag("hyperlink"):
            link_text = "".join(_render_run(run) for run in child.findall(_word_tag("r")))
            relation_id = child.get(f"{{{REL_NS}}}id")
            target = relationships.get(relation_id or "")
            if target:
                fragments.append(f'<a href="{html.escape(target, quote=True)}">{link_text}</a>')
            else:
                fragments.append(link_text)
        elif child.tag == _word_tag("sdt"):
            for run in child.iter(_word_tag("r")):
                fragments.append(_render_run(run))
        elif child.tag == _word_tag("pPr"):
            continue
        else:
            # Mantiene una descrizione utile per immagini quando Word la fornisce.
            if child.tag in (_word_tag("drawing"), _word_tag("pict")):
                description = next((item.get("descr") or item.get("title") for item in child.iter() if item.get("descr") or item.get("title")), "immagine")
                fragments.append(f"[Immagine: {html.escape(description)}]")

    paragraph_properties = paragraph.find(_word_tag("pPr"))
    style_name = ""
    numbering = None
    if paragraph_properties is not None:
        style = paragraph_properties.find(_word_tag("pStyle"))
        if style is not None:
            style_name = style.get(_word_tag("val"), "")
        num_properties = paragraph_properties.find(_word_tag("numPr"))
        if num_properties is not None and num_properties.find(_word_tag("numId")) is not None:
            numbering = num_properties.find(_word_tag("numId")).get(_word_tag("val"), "")

    content = "".join(fragments).strip()
    if not content:
        return ""
    heading = re.search(r"heading\s*([1-6])", style_name, re.IGNORECASE)
    if heading:
        level = int(heading.group(1))
        return f"<h{level}>{content}</h{level}>"
    if style_name.lower() in {"title", "titolo"}:
        return f"<h1>{content}</h1>"
    if style_name.lower() in {"subtitle", "sottotitolo"}:
        return f"<h2>{content}</h2>"
    if numbering and numbering != "0":
        return f"<li>{content}</li>"
    return f"<p>{content}</p>"


def _table_html(table: ElementTree.Element, relationships: dict[str, str]) -> str:
    """Rende una tabella DOCX in HTML, formato supportato da markdownify."""
    rows: list[str] = []
    for row in table.findall(_word_tag("tr")):
        cells: list[str] = []
        for cell in row.findall(_word_tag("tc")):
            cell_parts = [
                _paragraph_html(paragraph, relationships)
                for paragraph in cell.findall(_word_tag("p"))
            ]
            cell_content = "<br>".join(part[3:-4] if part.startswith("<p>") and part.endswith("</p>") else part for part in cell_parts if part)
            cells.append(f"<td>{cell_content}</td>")
        if cells:
            rows.append("<tr>" + "".join(cells) + "</tr>")
    return "<table>" + "".join(rows) + "</table>" if rows else ""


def convert_docx_fallback(path: Path) -> str:
    """Converte OOXML essenziale se MarkItDown non ha il modulo DOCX opzionale."""
    try:
        with zipfile.ZipFile(path) as archive:
            document = ElementTree.fromstring(archive.read("word/document.xml"))
            relationships = _read_relationships(archive)
            body = document.find(f"{_word_tag('body')}")
            if body is None:
                raise ValueError("Il documento Word non contiene un corpo leggibile.")
            blocks: list[str] = []
            list_items: list[str] = []

            def flush_list() -> None:
                if list_items:
                    blocks.append("<ul>" + "".join(list_items) + "</ul>")
                    list_items.clear()

            for element in body:
                if element.tag == _word_tag("p"):
                    block = _paragraph_html(element, relationships)
                    if block.startswith("<li>"):
                        list_items.append(block)
                    else:
                        flush_list()
                        if block:
                            blocks.append(block)
                elif element.tag == _word_tag("tbl"):
                    flush_list()
                    table = _table_html(element, relationships)
                    if table:
                        blocks.append(table)
            flush_list()

        converted = markdownify("\n".join(blocks), heading_style="ATX", bullets="-")
        converted = converted.replace("\r\n", "\n").strip()
        if not converted:
            raise ValueError("Il documento Word non contiene testo convertibile.")
        return converted + "\n"
    except (KeyError, zipfile.BadZipFile, ElementTree.ParseError) as exc:
        raise ValueError("Il documento è corrotto o non contiene OOXML leggibile.") from exc


def convert_docx(content: bytes, original_name: str) -> tuple[str, str, str]:
    """Valida, identifica e converte un DOCX; ritorna testo, motore e tipo."""
    validate_docx(content)
    identification = Magika().identify_bytes(content)
    detected_type = identification.output.label
    detected_mime = identification.output.mime_type
    if detected_type != "docx" or detected_mime != DOCX_MIME:
        raise ValueError(
            f"Magika identifica il file come {detected_type} ({detected_mime}), non come DOCX."
        )

    with tempfile.TemporaryDirectory(prefix="docx_markdown_") as temp_dir:
        temp_path = Path(temp_dir) / "documento.docx"
        temp_path.write_bytes(content)
        try:
            result = MarkItDown().convert(temp_path)
            markdown = result.text_content
            engine = "MarkItDown"
        except Exception as exc:
            message = str(exc).lower()
            missing_docx_extra = (
                "missingdependencyexception" in type(exc).__name__.lower()
                or "dependencies needed to read .docx" in message
                or "optional dependency [docx]" in message
            )
            if not missing_docx_extra:
                raise RuntimeError(f"MarkItDown non è riuscito a convertire il documento: {exc}") from exc
            markdown = convert_docx_fallback(temp_path)
            engine = "Fallback DOCX locale"

    markdown = markdown.replace("\r\n", "\n").strip()
    if not markdown:
        raise ValueError("La conversione non ha prodotto testo Markdown.")
    safe_stem = Path(original_name.replace("\\", "/")).stem
    safe_stem = re.sub(r"[^\w.-]+", "_", safe_stem, flags=re.UNICODE).strip("._") or "documento"
    return markdown + "\n", engine, f"{safe_stem}.md"


def convert_markdown(content: bytes, original_name: str) -> tuple[str, str, str]:
    """Legge Markdown UTF-8 in memoria e prepara il nome sicuro per il download."""
    if not content:
        raise ValueError("Il file Markdown è vuoto.")
    if len(content) > MAX_UPLOAD_BYTES:
        raise ValueError("Il file supera il limite consentito di 20 MB.")
    try:
        markdown = content.decode("utf-8-sig").replace("\r\n", "\n").replace("\r", "\n")
    except UnicodeDecodeError as exc:
        raise ValueError("Il file Markdown deve essere codificato in UTF-8.") from exc
    if not markdown.strip():
        raise ValueError("Il file Markdown è vuoto.")
    safe_stem = Path(original_name.replace("\\", "/")).stem
    safe_stem = re.sub(r"[^\w.-]+", "_", safe_stem, flags=re.UNICODE).strip("._") or "documento"
    return markdown if markdown.endswith("\n") else markdown + "\n", "Markdown", f"{safe_stem}.md"


def _read_upload(body: bytes, content_type: str) -> tuple[str, bytes]:
    """Legge il campo file dal multipart senza alterare i byte del documento."""
    if "multipart/form-data" not in content_type.lower():
        raise ValueError("La richiesta deve contenere un upload multipart/form-data.")
    message = BytesParser(policy=email.policy.default).parsebytes(
        f"MIME-Version: 1.0\r\nContent-Type: {content_type}\r\n\r\n".encode("ascii", "strict") + body
    )
    if not message.is_multipart():
        raise ValueError("Upload multipart non valido o boundary mancante.")
    for part in message.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        if part.get_param("name", header="content-disposition") != "file":
            continue
        filename = part.get_filename() or "documento.docx"
        payload = part.get_payload(decode=True) or b""
        return filename, payload
    raise ValueError("Nessun file trovato nella richiesta.")


def _send_json(handler: BaseHTTPRequestHandler, status: int, payload: dict) -> None:
    """Invia una risposta JSON con dimensione e intestazioni corrette."""
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.send_header("X-Content-Type-Options", "nosniff")
    handler.end_headers()
    handler.wfile.write(data)


class MarkdownHTTPRequestHandler(BaseHTTPRequestHandler):
    """Gestisce la pagina principale e l'API locale di conversione."""

    def log_message(self, format: str, *args: object) -> None:
        return

    def do_GET(self) -> None:
        """Serve l'interfaccia o segnala un percorso inesistente."""
        path = urllib.parse.urlparse(self.path).path
        if path not in ("/", "/index.html"):
            self.send_error(404)
            return
        data = HTML_PAGE.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self) -> None:
        """Riceve un DOCX o Markdown e restituisce il contenuto Markdown."""
        if urllib.parse.urlparse(self.path).path != "/api/convert":
            _send_json(self, 404, {"error": "Endpoint non trovato."})
            return
        origin = self.headers.get("Origin")
        if origin:
            expected_origin = f"http://{self.headers.get('Host', '')}"
            if origin.rstrip("/") != expected_origin:
                _send_json(self, 403, {"error": "Origine della richiesta non consentita."})
                return
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            _send_json(self, 400, {"error": "Dimensione della richiesta non valida."})
            return
        if content_length <= 0:
            _send_json(self, 400, {"error": "La richiesta non contiene un file."})
            return
        if content_length > MAX_UPLOAD_BYTES + 64 * 1024:
            _send_json(self, 413, {"error": "Il file supera il limite consentito di 20 MB."})
            return

        try:
            body = self.rfile.read(content_length)
            filename, content = _read_upload(body, self.headers.get("Content-Type", ""))
            if len(content) > MAX_UPLOAD_BYTES:
                raise ValueError("Il file supera il limite consentito di 20 MB.")
            if filename.lower().endswith(".docx"):
                markdown, engine, download_name = convert_docx(content, filename)
                detected_type = "Microsoft Word DOCX"
            elif filename.lower().endswith(".md"):
                markdown, engine, download_name = convert_markdown(content, filename)
                detected_type = "Markdown UTF-8"
            else:
                raise ValueError("Sono consentiti solo file con estensione .docx o .md.")
            _send_json(
                self,
                200,
                {
                    "markdown": markdown,
                    "engine": engine,
                    "download_name": download_name,
                    "detected_type": detected_type,
                },
            )
        except ValueError as exc:
            _send_json(self, 400, {"error": str(exc)})
        except Exception as exc:
            print(f"[ERRORE] Conversione DOCX non riuscita: {exc}", flush=True)
            _send_json(self, 422, {"error": "Conversione non riuscita. Verifica che il documento Word non sia danneggiato."})


def run_server(host: str = HOST, port: int = PORT) -> None:
    """Avvia il server locale e apre la pagina nel browser predefinito."""
    server = ThreadingHTTPServer((host, port), MarkdownHTTPRequestHandler)
    server.daemon_threads = True
    url = f"http://{host}:{port}/"
    print("=" * 62)
    print("Word in Markdown - Web app locale avviata")
    print(f"Indirizzo: {url}")
    print("Premi Ctrl+C per arrestare il server.")
    print("=" * 62, flush=True)
    threading.Thread(target=lambda: (time.sleep(0.8), webbrowser.open(url)), daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nArresto del server locale...", flush=True)
    finally:
        server.server_close()


if __name__ == "__main__":
    run_server()
