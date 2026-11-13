#!/usr/bin/env python3
"""Подготовка набора документов одного пациента для файлового режима аудита.

Читает .docx, .pdf (с текстовым слоем, нужен pdftotext), .txt, .md из папки или списка файлов,
закрывает типовые идентификаторы (ИИН, телефоны, почта, номера документов), убирает номера страниц,
декоративные линии и длинные дословные повторы абзацев (заменяет пометкой), пишет:
  <out>/doc_NN.txt   текст каждого документа под строкой «=== ДОКУМЕНТ N ===»
  <out>/all.txt      все документы подряд
  <out>/manifest.md  номер, размер, замечания (имена исходных файлов только здесь, в аудит не передаются)
ФИО, даты рождения, адреса и места работы автоматически НЕ закрываются: проверить текст вручную.

Запуск: python3 tools/prepare_case.py <папка или файлы> --out cases/<имя>/text
Папка cases/ в .gitignore: реальные данные в репозиторий не коммитить.
"""
import argparse, os, re, subprocess, sys, zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def runs_text(node):
    out = []
    for c in node:
        t = c.tag
        if t == W + "t": out.append(c.text or "")
        elif t == W + "tab": out.append("\t")
        elif t in (W + "br", W + "cr"): out.append("\n")
        elif t in (W + "pPr", W + "rPr", W + "delText", W + "instrText", W + "del", W + "moveFrom"): continue
        else: out.append(runs_text(c))
    return "".join(out)

def block_lines(node, out):
    for c in node:
        t = c.tag
        if t == W + "p": out.append(runs_text(c).rstrip())
        elif t == W + "tbl":
            for tr in c.findall(W + "tr"):
                cells = []
                for tc in tr.findall(W + "tc"):
                    inner = []; block_lines(tc, inner)
                    cells.append(re.sub(r"\s+", " ", " ".join(inner)).strip())
                line = " | ".join(cells).strip()
                if line.replace("|", "").strip(): out.append(line)
        elif t == W + "sdt":
            sc = c.find(W + "sdtContent")
            if sc is not None: block_lines(sc, out)
        elif t in (W + "ins", W + "customXml", W + "smartTag"): block_lines(c, out)

def read_docx(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    lines = []; block_lines(body, lines)
    return "\n".join(lines)

def read_pdf(path):
    try:
        return subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True, check=True).stdout
    except Exception as e:
        raise RuntimeError("не удалось прочитать PDF (нужен pdftotext): %s" % e)

def read_any(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx": return read_docx(path)
    if ext == ".pdf": return read_pdf(path)
    if ext in (".txt", ".md"): return open(path, encoding="utf-8", errors="replace").read()
    raise RuntimeError("формат %s не читается (для .doc сохраните как .docx)" % ext)

def mask(t):
    c = {"ИИН": 0, "телефоны": 0, "почта": 0, "документы": 0}
    def sub(rx, repl, key, flags=0):
        nonlocal t
        def f(m): c[key] += 1; return repl(m) if callable(repl) else repl
        t = re.sub(rx, f, t, flags=flags)
    sub(r"(?<!\d)\d{12}(?!\d)", "[ИИН закрыт]", "ИИН")
    sub(r"(?:\+7|8)[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}", "[телефон закрыт]", "телефоны")
    sub(r"[\w.+\-]+@[\w\-]+\.[\w.\-]+", "[почта закрыта]", "почта")
    sub(r"(Удостоверение личности|Паспорт)[^\n]{0,40}\d{6,}", lambda m: m.group(1) + " [номер закрыт]", "документы", re.I)
    return t, c

def clean(t):
    seen, st, out = set(), {"страницы": 0, "линии": 0, "повторы": 0}, []
    for line in t.replace("\r\n", "\n").replace("\r", "\n").replace(" ", " ").split("\n"):
        l = line.rstrip(); core = l.strip()
        if re.fullmatch(r"(?i)(стр(аница|\.)?\s*)\d+(\s*(из|/)\s*\d+)?", core) or re.fullmatch(r"\d+\s*(из|/)\s*\d+", core):
            st["страницы"] += 1; continue
        if re.fullmatch(r"[-–—_.=*~ ]{4,}", core): st["линии"] += 1; continue
        if len(core) >= 140:
            key = re.sub(r"\s+", " ", core).lower()
            if key in seen:
                st["повторы"] += 1; out.append("[дословный повтор записи выше: «%s…»]" % core[:50]); continue
            seen.add(key)
        out.append(l)
    return re.sub(r"\n{3,}", "\n\n", "\n".join(out)).strip(), st

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="+", help="папка или файлы документов одного пациента")
    ap.add_argument("--out", required=True, help="папка для результата")
    a = ap.parse_args()
    files = []
    for p in a.inputs:
        if os.path.isdir(p):
            files += sorted(os.path.join(p, f) for f in os.listdir(p) if not f.startswith("."))
        else: files.append(p)
    os.makedirs(a.out, exist_ok=True)
    manifest = ["# Состав материалов (локально; имена файлов в аудит не передавать)", "",
                "| № | Файл | Знаков | Закрыто автоматически | Убрано | Замечание |", "|---|---|---|---|---|---|"]
    alltxt = []; n = 0
    for f in files:
        try:
            raw = read_any(f); note = ""
        except Exception as e:
            manifest.append("| - | %s | 0 | - | - | %s |" % (os.path.basename(f), e)); continue
        n += 1
        t, mc = mask(raw); t, cs = clean(t)
        if len(re.sub(r"\s", "", t)) < 200: note = "почти нет текста (скан или картинки): распознать или вставить вручную"
        block = "=== ДОКУМЕНТ %d ===\n%s\n" % (n, t)
        open(os.path.join(a.out, "doc_%02d.txt" % n), "w", encoding="utf-8").write(block)
        alltxt.append(block)
        manifest.append("| %d | %s | %d | %s | %s | %s |" % (n, os.path.basename(f), len(t),
            ", ".join("%s %d" % (k, v) for k, v in mc.items() if v) or "-",
            ", ".join("%s %d" % (k, v) for k, v in cs.items() if v) or "-", note or "-"))
    open(os.path.join(a.out, "all.txt"), "w", encoding="utf-8").write("\n".join(alltxt))
    manifest += ["", "Всего документов: %d, знаков: %d." % (n, sum(len(x) for x in alltxt)),
                 "ФИО, даты рождения, адреса, места работы автоматически не закрываются: проверить текст вручную до аудита."]
    open(os.path.join(a.out, "manifest.md"), "w", encoding="utf-8").write("\n".join(manifest) + "\n")
    print("\n".join(manifest))

if __name__ == "__main__":
    main()
