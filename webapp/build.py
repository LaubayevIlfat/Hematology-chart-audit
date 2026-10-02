"""Собирает webapp/page.html: вшивает текст навыка и шаблона заключения в template.html.

Запуск из корня репозитория: python3 webapp/build.py
Страница публикуется как Claude Artifact с capabilities {"sample": {}, "downloads": true}.
При изменении правил навыка (SKILL.md, references/, шаблон) страницу нужно пересобрать и опубликовать заново.
"""
import json, os, re, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
base = os.path.join(ROOT, ".claude", "skills", "hematology-chart-audit")

with zipfile.ZipFile(os.path.join(ROOT, "templates", "Shablon_zaklyucheniya_SPPVE.docx")) as z:
    xml = z.read("word/document.xml").decode("utf8")
paras = []
for p in re.findall(r"<w:p[ >].*?</w:p>", xml, flags=re.S):
    t = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p, flags=re.S))
    if t.strip():
        paras.append(t.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">"))

parts = {"SKILL.md": open(os.path.join(base, "SKILL.md"), encoding="utf8").read()}
for f in sorted(os.listdir(os.path.join(base, "references"))):
    parts["references/" + f] = open(os.path.join(base, "references", f), encoding="utf8").read()
parts["template"] = "ШАБЛОН ЗАКЛЮЧЕНИЯ СППВЭ (текст из Shablon_zaklyucheniya_SPPVE.docx):\n" + "\n".join(paras)

data = json.dumps({"parts": parts}, ensure_ascii=False).replace("</", "<\\/")
tpl = open(os.path.join(ROOT, "webapp", "template.html"), encoding="utf8").read()
assert "__SKILL_JSON__" in tpl
out = os.path.join(ROOT, "webapp", "page.html")
open(out, "w", encoding="utf8").write(tpl.replace("__SKILL_JSON__", data))
print("written", out, os.path.getsize(out), "bytes")
