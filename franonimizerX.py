import os
import re
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, scrolledtext
import pymupdf
import docx
import spacy
import Vision
from Foundation import NSURL
from PIL import Image, ImageDraw

OUTPUT_DIR = os.path.expanduser("~/Documents/anonymized")
os.makedirs(OUTPUT_DIR, exist_ok=True)

nlp = spacy.load("sv_core_news_sm")

PN_REGEX = re.compile(r'\b((?:19|20)?\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])[-+]?)\d{4}\b')
EMAIL_REGEX = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
PHONE_REGEX = re.compile(r'\b(?:\+46|0046|0)[\s-]?\d{1,3}[\s-]?\d{2,3}[\s-]?\d{2}[\s-]?\d{2,3}\b')
ADDRESS_REGEX = re.compile(
    r'\b[A-ZÅÄÖ][a-zåäöA-ZÅÄÖ]*(?:\s+[a-zåäöA-ZÅÄÖ]+)?(?:gatan|gata|vägen|väg|gränd|allén?|torget|torg|stigen|stig|backen|platsen)\s*\d*[A-Za-z]?(?:\s*(?:lgh|tr|vån)\s*\d+)?\b',
    re.IGNORECASE
)

TEXTS = {
    "sv": {
        "desc": "Välj PDF-, Word- (.docx), text- eller bildfiler.\nOriginalfilerna lämnas orörda.",
        "btn_select": "Välj filer...",
        "btn_open": "Öppna utdatamapp",
        "dialog_title": "Välj filer att anonymisera",
        "ft_supported": "Filer som stöds",
        "ft_all": "Alla filer",
        "processing": "Bearbetar: {name}...",
        "done": "Klar: {name}",
        "unsupported": "Formatet stöds inte: {name}",
        "not_file": "Inte en fil: {name}",
        "error": "Fel ({name}): {err}",
    },
    "en": {
        "desc": "Select PDF, Word (.docx), text, or image files.\nOriginal files remain untouched.",
        "btn_select": "Select files...",
        "btn_open": "Open output folder",
        "dialog_title": "Select files to anonymize",
        "ft_supported": "Supported files",
        "ft_all": "All files",
        "processing": "Processing: {name}...",
        "done": "Done: {name}",
        "unsupported": "Unsupported format: {name}",
        "not_file": "Not a file: {name}",
        "error": "Error ({name}): {err}",
    },
    "hu": {
        "desc": "Válassz ki PDF, Word (.docx), szöveg- vagy képfájlokat.\nAz eredeti fájlok érintetlenek maradnak.",
        "btn_select": "Fájlok kiválasztása...",
        "btn_open": "Kimeneti mappa megnyitása",
        "dialog_title": "Fájlok kiválasztása",
        "ft_supported": "Támogatott fájlok",
        "ft_all": "Minden fájl",
        "processing": "Feldolgozás: {name}...",
        "done": "Kész: {name}",
        "unsupported": "Nem támogatott formátum: {name}",
        "not_file": "Nem fájl: {name}",
        "error": "Hiba ({name}): {err}",
    }
}

def name_to_initials(name):
    parts = [p for p in name.strip().split() if p]
    return " ".join([f"{p[0].upper()}." for p in parts])

def mask_text(text):
    if not text:
        return text
    text = PN_REGEX.sub(r'\1XXXX', text)
    text = EMAIL_REGEX.sub('[E-POST]', text)
    text = PHONE_REGEX.sub('[TELEFON]', text)
    text = ADDRESS_REGEX.sub('[ADRESS]', text)
    doc = nlp(text)
    for ent in sorted(doc.ents, key=lambda x: x.start_char, reverse=True):
        if ent.label_ in ("PRS", "PER"):
            initials = name_to_initials(ent.text)
            text = text[:ent.start_char] + initials + text[ent.end_char:]
        elif ent.label_ in ("LOC", "GPE"):
            text = text[:ent.start_char] + '[PLATS]' + text[ent.end_char:]
    return text

def process_txt(input_path, output_path):
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(mask_text(content))

def process_docx(input_path, output_path):
    doc = docx.Document(input_path)
    for para in doc.paragraphs:
        if para.text.strip():
            para.text = mask_text(para.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    cell.text = mask_text(cell.text)
    doc.save(output_path)

def process_pdf(input_path, output_path):
    doc = pymupdf.open(input_path)
    for page in doc:
        text = page.get_text("text")
        for match in PN_REGEX.finditer(text):
            full_pn = match.group(0)
            masked_pn = match.group(1) + "XXXX"
            for rect in page.search_for(full_pn):
                page.add_redact_annot(rect, text=masked_pn, fill=(1, 1, 1), text_color=(0, 0, 0))
        for pattern, label in [
            (EMAIL_REGEX, "[E-POST]"),
            (PHONE_REGEX, "[TELEFON]"),
            (ADDRESS_REGEX, "[ADRESS]")
        ]:
            for match in pattern.finditer(text):
                for rect in page.search_for(match.group(0)):
                    page.add_redact_annot(rect, text=label, fill=(0, 0, 0), text_color=(1, 1, 1))
        spacy_doc = nlp(text)
        for ent in spacy_doc.ents:
            if ent.label_ in ("PRS", "PER") and len(ent.text.strip()) > 2:
                initials = name_to_initials(ent.text)
                for rect in page.search_for(ent.text):
                    page.add_redact_annot(rect, text=initials, fill=(1, 1, 1), text_color=(0, 0, 0))
            elif ent.label_ in ("LOC", "GPE") and len(ent.text.strip()) > 2:
                for rect in page.search_for(ent.text):
                    page.add_redact_annot(rect, text="[PLATS]", fill=(0, 0, 0), text_color=(1, 1, 1))
        page.apply_redactions()
    doc.save(output_path)

def process_image(input_path, output_path):
    img = Image.open(input_path)
    w, h = img.size
    draw = ImageDraw.Draw(img)
    url = NSURL.fileURLWithPath_(os.path.abspath(input_path))
    handler = Vision.VNImageRequestHandler.alloc().initWithURL_options_(url, None)
    request = Vision.VNRecognizeTextRequest.alloc().init()
    request.setRecognitionLanguages_(["sv-SE", "en-US"])
    request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
    handler.performRequests_error_([request], None)
    for obs in request.results() or []:
        candidates = obs.topCandidates_(1)
        if not candidates:
            continue
        line_text = candidates[0].string()
        if PN_REGEX.search(line_text) or EMAIL_REGEX.search(line_text) or PHONE_REGEX.search(line_text) or ADDRESS_REGEX.search(line_text):
            bbox = obs.boundingBox()
            x = bbox.origin.x * w
            y = (1.0 - bbox.origin.y - bbox.size.height) * h
            draw.rectangle([x, y, x + bbox.size.width * w, y + bbox.size.height * h], fill="black")
    img.save(output_path)

def handle_file(file_path, lang="sv"):
    t = TEXTS[lang]
    if not os.path.isfile(file_path):
        return t["not_file"].format(name=file_path)
    base = os.path.basename(file_path)
    name, ext = os.path.splitext(base)
    ext_lower = ext.lower()
    output_path = os.path.join(OUTPUT_DIR, f"{name}_GDPR{ext}")
    try:
        if ext_lower == ".pdf":
            process_pdf(file_path, output_path)
        elif ext_lower == ".docx":
            process_docx(file_path, output_path)
        elif ext_lower in (".txt", ".csv", ".md"):
            process_txt(file_path, output_path)
        elif ext_lower in (".png", ".jpg", ".jpeg"):
            process_image(file_path, output_path)
        else:
            return t["unsupported"].format(name=base)
        return t["done"].format(name=f"{name}_GDPR{ext}")
    except Exception as e:
        return t["error"].format(name=base, err=e)

def start_gui():
    root = tk.Tk()
    root.title("Franonimizer")
    root.geometry("560x390")
    current_lang = tk.StringVar(value="sv")
    lang_frame = tk.Frame(root, pady=8)
    lang_frame.pack()

    def update_ui():
        t = TEXTS[current_lang.get()]
        desc_label.config(text=t["desc"])
        btn_select.config(text=t["btn_select"])
        btn_open.config(text=t["btn_open"])

    for code, label_text in [("sv", "Svenska"), ("en", "English"), ("hu", "Magyar")]:
        tk.Radiobutton(lang_frame, text=label_text, variable=current_lang, value=code, command=update_ui, padx=8).pack(side=tk.LEFT)

    desc_label = tk.Label(root, text="", pady=8)
    desc_label.pack()
    log_box = scrolledtext.ScrolledText(root, width=64, height=12, state="disabled")
    log_box.pack(padx=12, pady=6)

    def write_log(msg):
        log_box.config(state="normal")
        log_box.insert(tk.END, msg + "\n")
        log_box.see(tk.END)
        log_box.config(state="disabled")
        root.update()

    def select_files():
        lang = current_lang.get()
        t = TEXTS[lang]
        paths = filedialog.askopenfilenames(
            title=t["dialog_title"],
            filetypes=[(t["ft_supported"], "*.pdf *.docx *.txt *.csv *.md *.png *.jpg *.jpeg"), (t["ft_all"], "*.*")]
        )
        for p in paths:
            write_log(t["processing"].format(name=os.path.basename(p)))
            write_log(handle_file(p, lang=lang))

    btn_frame = tk.Frame(root, pady=10)
    btn_frame.pack()
    btn_select = tk.Button(btn_frame, text="", command=select_files, padx=12, pady=6)
    btn_select.pack(side=tk.LEFT, padx=8)
    btn_open = tk.Button(btn_frame, text="", command=lambda: subprocess.run(["open", OUTPUT_DIR]), padx=12, pady=6)
    btn_open.pack(side=tk.LEFT, padx=8)
    update_ui()
    root.mainloop()

if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("-psn_")]
    if args:
        for path in args:
            print(handle_file(path))
    else:
        start_gui()
