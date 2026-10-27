# -*- coding: utf-8 -*-
"""
Bot de Telegram para la boveda de Obsidian (version 2, disenada por el usuario).

Capacidades:
- Texto, notas de voz, fotos del cuaderno/pizarra, PDF, Word, enlaces web y YouTube.
- Clasifica el contenido (teoria / esquema / ejercicios / trabajos), detecta
  asignatura y tema, y archiva siguiendo la jerarquia y numeracion de la boveda.
- Crear notas: directo. Crear carpeta, mover, renombrar, editar y /organizar:
  siempre pide confirmacion en el chat. Nunca borra nada.
- RAG: /pregunta responde desde tus propias notas citando el archivo fuente.
- Sube todos los cambios a GitHub (git add/commit/pull/push).

Dependencias (ver requirements.txt):
requests, beautifulsoup4, youtube-transcript-api
Los .docx se leen con librerias estandar de Python (zipfile + XML),
asi no hace falta python-docx.
"""

import base64
import io
import json
import os
import re
import subprocess
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import requests
from bs4 import BeautifulSoup

try:
    from youtube_transcript_api import YouTubeTranscriptApi
except Exception:
    YouTubeTranscriptApi = None

try:
    import config
except ImportError:
    config = None


def _cfg(entrorno, nombre_config, por_defecto=None):
    valor = os.environ.get(entrorno)
    if valor is not None and valor != "":
        return valor
    if config is not None and hasattr(config, nombre_config):
        return getattr(config, nombre_config)
    return por_defecto


BOT_TOKEN = _cfg("BOT_TOKEN", "BOT_TOKEN")
GEMINI_KEY = _cfg("GEMINI_API_KEY", "GEMINI_API_KEY")
GEMINI_MODEL = _cfg("GEMINI_MODEL", "GEMINI_MODEL", "gemini-3.5-flash-lite")
ALLOWED_CHAT_ID = int(_cfg("ALLOWED_CHAT_ID", "ALLOWED_CHAT_ID", 0) or 0)
VAULT = Path(_cfg("VAULT_PATH", "VAULT_PATH", ".")).resolve()
TUTOR_MODE = str(_cfg("TUTOR_MODE", "TUTOR_MODE", False)).lower() == "true"
IGNORE_DIRS = getattr(config, "IGNORE_DIRS",
                      {".git", ".obsidian", ".trash", ".tmp", "node_modules"})
ATTACHMENTS_DIR = _cfg("ATTACHMENTS_DIR", "ATTACHMENTS_DIR", "Adjuntos")
# 0 = sin limite (uso local); en la nube se pone ~340 para el relevo
MAX_RUNTIME_MINUTES = float(_cfg("MAX_RUNTIME_MINUTES", "MAX_RUNTIME_MINUTES", 0) or 0)

TG_API = f"https://api.telegram.org/bot{BOT_TOKEN}"
GEMINI_URL = (
    f"https://generativelanguage.googleapis.com/v1beta/models/"
    f"{GEMINI_MODEL}:generateContent?key={GEMINI_KEY}"
)

# Acciones pendientes de confirmacion: chat_id -> diccionario con la accion
PENDING = {}

URL_RE = re.compile(r"https?://[^\s]+")
YOUTUBE_RE = re.compile(
    r"(?:youtube\.com/watch\?v=|youtu\.be/|youtube\.com/shorts/)([\w-]{6,})")


# ----------------------------------------------------------------------
# Telegram
# ----------------------------------------------------------------------
def tg(method, payload=None):
    r = requests.post(f"{TG_API}/{method}", json=payload or {}, timeout=90)
    return r.json()


def send_message(chat_id, text):
    for i in range(0, len(text), 3800):
        tg("sendMessage", {
            "chat_id": chat_id,
            "text": text[i:i + 3800],
            "disable_web_page_preview": True,
        })


def send_document(chat_id, file_path):
    with open(file_path, "rb") as f:
        requests.post(
            f"{TG_API}/sendDocument",
            data={"chat_id": chat_id},
            files={"document": (Path(file_path).name, f)},
            timeout=180,
        )


def download_telegram_file(file_id):
    info = tg("getFile", {"file_id": file_id})
    if not info.get("ok"):
        raise RuntimeError("Telegram no deja descargar el archivo "
                           "(quizas supera 20 MB).")
    path = info["result"]["file_path"]
    r = requests.get(f"https://api.telegram.org/file/bot{BOT_TOKEN}/{path}",
                     timeout=180)
    r.raise_for_status()
    return r.content


# ----------------------------------------------------------------------
# Boveda
# ----------------------------------------------------------------------
def md_files():
    out = []
    for root, dirs, files in os.walk(VAULT):
        dirs[:] = [d for d in dirs if d not in IGNORE_DIRS]
        for name in files:
            if name.lower().endswith(".md"):
                out.append(Path(root) / name)
    return sorted(out)


def all_folders():
    out = set()
    for root, dirs, _ in os.walk(VAULT):
        dirs[:] = [d for d in dirs if d not in IGNORE_DIRS]
        for d in dirs:
            out.add(rel(Path(root) / d))
    return sorted(out)


def rel(p):
    return str(Path(p).resolve().relative_to(VAULT)).replace("\\", "/")


def vault_tree():
    lineas = [rel(f) for f in md_files()]
    return "\n".join(lineas) if lineas else "(la boveda esta vacia)"


def find_note(query):
    q = query.lower().strip().removesuffix(".md")
    candidatos = []
    for f in md_files():
        ruta = rel(f)
        if q == f.stem.lower() or q in ruta.lower():
            candidatos.append(f)
    return min(candidatos, key=lambda f: len(rel(f))) if candidatos else None


def ruta_segura(ruta_texto):
    ruta_texto = ruta_texto.strip().replace("\\", "/").strip("/")
    destino = (VAULT / ruta_texto).resolve()
    if not str(destino).startswith(str(VAULT)):
        raise ValueError("Ruta fuera de la boveda")
    if any(part in IGNORE_DIRS for part in Path(ruta_texto).parts):
        raise ValueError("Carpeta protegida, no se puede tocar")
    return destino


def safe_path(folder, filename):
    folder = (folder or "").strip().replace("\\", "/").strip("/")
    filename = Path(filename).name.strip()
    if not filename.lower().endswith(".md"):
        filename += ".md"
    return ruta_segura(f"{folder}/{filename}" if folder else filename)


def crear_carpeta(ruta_carpeta):
    destino = ruta_segura(ruta_carpeta)
    destino.mkdir(parents=True, exist_ok=True)
    return rel(destino)


def mover_nota(origen_texto, destino_texto):
    origen = ruta_segura(origen_texto)
    if not origen.exists():
        encontrado = find_note(origen_texto)
        if not encontrado:
            raise FileNotFoundError(f"No existe: {origen_texto}")
        origen = encontrado

    destino = ruta_segura(destino_texto)
    if destino.is_dir() or destino_texto.endswith("/"):
        destino = destino / origen.name
    elif destino.suffix.lower() != ".md":
        destino = destino.with_suffix(".md")

    if destino.exists():
        raise FileExistsError(f"Ya existe un archivo en {rel(destino)}")
    destino.parent.mkdir(parents=True, exist_ok=True)
    origen.rename(destino)
    return rel(origen), rel(destino)


# ----------------------------------------------------------------------
# Busqueda local (RAG sencillo, sin servicios externos)
# ----------------------------------------------------------------------
STOPWORDS = {
    "que", "qué", "como", "cómo", "donde", "dónde", "cuando", "cuándo",
    "este", "esta", "estos", "estas", "esto", "aque", "aquell", "sobre",
    "para", "porque", "porqué", "otro", "otra", "pero", "entre", "hacia",
    "hasta", "desde", "tambien", "también", "tiene", "tienen", "puede",
    "the", "and", "for", "with", "that", "this", "are", "was", "were",
    "have", "has", "from", "into", "your", "you", "all", "not", "but",
}


def _palabras(texto):
    return [p for p in re.findall(r"\w[\w'’-]{2,}", texto.lower())
            if p not in STOPWORDS and len(p) > 3]


def buscar_fragmentos(consulta, top_n=5):
    terminos = _palabras(consulta)
    if not terminos:
        return []
    puntuadas = []
    for f in md_files():
        try:
            lineas = f.read_text(encoding="utf-8", errors="ignore").splitlines()
        except Exception:
            continue
        puntos = 0
        mejor_linea = ""
        for linea in lineas:
            low = linea.lower()
            golpe = sum(1 for t in terminos if t in low)
            if golpe > puntos:
                puntos = golpe
                mejor_linea = linea.strip()
        if puntos:
            puntuadas.append((puntos, f, mejor_linea[:200]))
    puntuadas.sort(key=lambda x: x[0], reverse=True)
    return puntuadas[:top_n]


# ----------------------------------------------------------------------
# Git
# ----------------------------------------------------------------------
def git(*args):
    res = subprocess.run(
        ["git", "-C", str(VAULT), *args],
        capture_output=True, text=True, timeout=180, encoding="utf-8",
        errors="replace",
    )
    return res.returncode, ((res.stdout or "") + (res.stderr or "")).strip()


def git_guardar(mensaje):
    code, out = git("add", "-A")
    if code != 0:
        return False, f"git add fallo:\n{out}"
    code, out = git("commit", "-m", mensaje)
    if code != 0 and "nothing to commit" not in out.lower():
        return False, f"git commit fallo:\n{out}"
    git("pull", "--no-rebase", "--no-edit")
    code, out = git("push")
    if code != 0:
        return False, ("Guardado en el PC, pero el push fallo:\n" + out +
                       "\nEl plugin Git de Obsidian lo subira luego.")
    return True, "Subido a GitHub: llegara al movil en la proxima sincronizacion."


# ----------------------------------------------------------------------
# Gemini
# ----------------------------------------------------------------------
def gemini(parts, json_mode=False):
    payload = {"contents": [{"parts": parts}]}
    if json_mode:
        payload["generationConfig"] = {"responseMimeType": "application/json"}
    r = requests.post(GEMINI_URL, json=payload, timeout=180)
    if r.status_code != 200:
        raise RuntimeError(f"Gemini responde {r.status_code}: {r.text[:400]}")
    data = r.json()
    return data["candidates"][0]["content"]["parts"][0]["text"]


def _parsear_json(crudo):
    texto = crudo.strip().removeprefix("```json").removeprefix("```")
    return json.loads(texto.removesuffix("```").strip())


def transcribir_audio(audio_bytes):
    return gemini([
        {"text": "Transcribe literalmente este audio (espanol o catalan). "
                 "Solo la transcripcion."},
        {"inline_data": {"mime_type": "audio/ogg",
                         "data": base64.b64encode(audio_bytes).decode()}},
    ]).strip()


def extraer_imagen(img_bytes, mime):
    prompt = """Analiza esta imagen (puede ser una foto de cuaderno, pizarra,
apunte a mano o enunciado). Devuelve JSON:
{"asignatura_estimada":"FiQ|Historia|Ingles|incierta",
 "tipo":"teoria|esquema|ejercicio|trabajo|otro",
 "tema_estimado":"texto corto",
 "markdown":"transcripcion fiel y limpia del contenido en markdown,
             manteniendo titulos, listas y formulas lo mejor posible"}
No inventes contenido que no se vea."""
    datos = _parsear_json(gemini([
        {"text": prompt},
        {"inline_data": {"mime_type": mime,
                         "data": base64.b64encode(img_bytes).decode()}},
    ], json_mode=True))
    return datos


def extraer_pdf(pdf_bytes):
    prompt = """Extrae y estructura el contenido de este PDF en markdown limpio.
Devuelve JSON:
{"asignatura_estimada":"FiQ|Historia|Ingles|incierta",
 "tipo":"teoria|esquema|ejercicio|trabajo|otro",
 "tema_estimado":"texto corto",
 "markdown":"contenido completo bien organizado con titulos y listas"}"""
    datos = _parsear_json(gemini([
        {"text": prompt},
        {"inline_data": {"mime_type": "application/pdf",
                         "data": base64.b64encode(pdf_bytes).decode()}},
    ], json_mode=True))
    return datos


def extraer_docx(docx_bytes, nombre_tmp):
    """Lee un .docx sin dependencias: es un ZIP con word/document.xml."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
            xml_bytes = z.read("word/document.xml")
    except Exception as e:
        raise RuntimeError(f"No es un .docx valido: {e}")
    root = ET.fromstring(xml_bytes)
    parrafos = []
    for p in root.iter(f"{{{w}}}p"):
        linea = "".join(t.text or "" for t in p.iter(f"{{{w}}}t")).strip()
        if linea:
            parrafos.append(linea)
    return {
        "asignatura_estimada": "incierta",
        "tipo": "otro",
        "tema_estimado": nombre_tmp.removesuffix(".docx"),
        "markdown": "\n\n".join(parrafos),
    }


def extraer_web(url):
    headers = {"User-Agent": "Mozilla/5.0 (compatible; ObsidianBot/1.0)"}
    r = requests.get(url, headers=headers, timeout=60)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header"]):
        tag.decompose()
    titulo = soup.title.get_text(strip=True) if soup.title else url
    texto = re.sub(r"\n{3,}", "\n\n", soup.get_text("\n", strip=True))
    return titulo, texto[:MAX_WEB_CHARS]


def extraer_youtube(url):
    if YouTubeTranscriptApi is None:
        raise RuntimeError("Falta youtube-transcript-api.")
    m = YOUTUBE_RE.search(url)
    if not m:
        raise ValueError("No reconozco el ID del video de YouTube.")
    video_id = m.group(1)
    try:
        # API de youtube-transcript-api 0.6.x (compatible con Python 3.15)
        segmentos = YouTubeTranscriptApi.get_transcript(
            video_id, languages=["es", "ca", "en"])
        texto = " ".join(seg["text"] for seg in segmentos)
    except Exception as e:
        raise RuntimeError(
            "No encuentro subtitulos para ese video (el canal no los tiene "
            f"o estan capados): {e}")
    return texto[:MAX_WEB_CHARS]


# ----------------------------------------------------------------------
# Clasificacion y archivo (IA)
# ----------------------------------------------------------------------
def prompt_sistema(contexto_extra=""):
    base = f"""Eres un gestor de archivos que convierte mensajes en notas de una
boveda Obsidian de un alumno de 4to de ESO. Tu trabajo principal es CREAR y
ORGANIZAR archivos, no charlar.

NOTAS ACTUALES:
{vault_tree()}

CARPETAS ACTUALES:
{chr(10).join(all_folders())}

REGLAS DE JERARQUIA:
- Raiz de asignaturas: "001- 4to ESO" con FiQ, Historia, Ingles.
- Secciones: 01- Indice, 02- Temario, 03- Trabajos, 04- Ejercicios.
- Numeracion jerarquica: mira los nombres existentes y continua el orden;
  el nombre lleva el prefijo numerico y el titulo del tema.
- Teoria -> 02- Temario. Ejercicios -> 04- Ejercicios. Trabajos -> 03- Trabajos.
- FiQ e Historia en catalan; Ingles en ingles. Tus respuestas, en espanol.
- No tocar .git, .obsidian, .trash. No anadir el campo "publish".
- Nunca proponer borrar archivos.

REGLA DE ORO:
- Si el mensaje pide crear, guardar, anotar, archivar, resumir, esquematizar
  o registrar algo, DEBES devolver una accion de archivo. NUNCA respondas solo
  con texto pidiendo mas informacion: decide tu la carpeta y el nombre.
- reply_only se usa UNICAMENTE para preguntas, dudas conceptuales o charla
  que no implica guardar nada.
- Si la carpeta destino no existe, devuelve propose_folder (no reply_only).
- Si NO tienes claro donde va (nota suelta, prueba, sin asignatura clara o
  contenido que podria ir a dos sitios), devuelve ask_folder con 2 o 3
  opciones de carpetas EXISTENTES. No adivines en casos dudosos.
- El campo "reply" de toda accion de archivo debe incluir un mini resumen
  de 2 o 3 lineas del contenido guardado, ademas de la ruta.

EJEMPLO:
Mensaje: "crea una nota de prueba tal cual en la boveda"
Respuesta correcta:
{{"action":"ask_folder","options":["","001- 4to ESO/Historia/04- Ejercicios"],
 "filename":"prueba.md","content":"# Prueba",
 "reply":"Nota de prueba sin asignatura; elige donde la guardo."}}

ACCIONES POSIBLES (devuelve solo JSON valido):
1 create_note:
  {{"action":"create_note","folder":"carpeta o vacio","filename":"....md",
    "content":"markdown completo","reply":"mini resumen + ruta"}}
2 append_note:
  {{"action":"append_note","target_note":"nombre o ruta",
    "content":"bloque a anadir","reply":"mensaje corto"}}
3 edit_note:
  {{"action":"edit_note","target_note":"...","old_text":"texto exacto actual",
    "new_text":"texto nuevo","reply":"que cambia"}}
4 move_note:
  {{"action":"move_note","source":"...","destination":"ruta completa",
    "reply":"mensaje corto"}}
5 propose_folder (carpeta nueva, pide confirmacion):
  {{"action":"propose_folder","folder":"carpeta nueva",
    "filename":"....md","content":"markdown","reply":"por que hace falta"}}
6 ask_folder (estas dudando; el alumno elige carpeta):
  {{"action":"ask_folder","options":["carpeta1","carpeta2"],
    "filename":"....md","content":"markdown completo","reply":"por que dudas"}}
7 reply_only (solo para dudas o conversacion sin guardar nada):
  {{"action":"reply_only","reply":"respuesta"}}
"""
    if TUTOR_MODE:
        base += """
REGLA DEL TUTOR: ante una tarea entera puntuable, no la entregues hecha;
devuelve reply_only con explicacion, esquema y preguntas guia.
"""
    if contexto_extra:
        base += "\nCONTEXTO ADICIONAL:\n" + contexto_extra
    return base


VERBOS_CREACION = (
    "crea", "crear", "guarda", "guardar", "anota", "apunta", "archiva",
    "registra", "resume", "resumen", "esquema", "ficha", "nota", "apunte",
    "apunta", "fes", "guarda'm", "crea'm", "anota",
)


def _ia_accion(texto_usuario, contexto=""):
    return _parsear_json(gemini([
        {"text": prompt_sistema(contexto)},
        {"text": "Mensaje del alumno:\n" + texto_usuario},
    ], json_mode=True))


def procesar_texto(texto_usuario, contexto="", chat_id=None,
                   forzar_accion=False, insertar=""):
    datos = _ia_accion(texto_usuario, contexto)

    # Si el alumno pidio guardar algo y la IA se puso a charlar, forzamos
    # una segunda respuesta mas estricta.
    bajo = texto_usuario.lower()
    pide_archivo = forzar_accion or any(v in bajo for v in VERBOS_CREACION)
    if pide_archivo and datos.get("action") == "reply_only":
        datos = _parsear_json(gemini([
            {"text": prompt_sistema(contexto)},
            {"text": "Mensaje del alumno:\n" + texto_usuario},
            {"text": "Tu respuesta anterior fue solo conversacion y es "
                     "incorrecta. El alumno pide guardar algo. Devuelve "
                     "OBLIGATORIAMENTE un JSON de accion create_note, "
                     "append_note, ask_folder o propose_folder. Nada de reply_only."},
        ], json_mode=True))

    # Insercion forzada (ej: imagen de una foto) al principio del contenido
    if insertar and datos.get("action") in (
            "create_note", "propose_folder", "ask_folder"):
        if insertar not in datos.get("content", ""):
            datos["content"] = insertar + "\n\n" + datos.get("content", "")

    if chat_id is not None:
        datos["_chat_id"] = chat_id
    return ejecutar_accion(datos, confirmar=False)


def responder_rag(consulta):
    fragmentos = buscar_fragmentos(consulta)
    if not fragmentos:
        return "No encuentro nada relacionado en tus notas."
    contexto = "\n\n".join(
        f"FUENTE: {rel(f)}\n{trozos}" for _, f, trozos in fragmentos)
    respuesta = gemini([
        {"text": "Responde la pregunta del alumno usando SOLO estos "
                 "fragmentos de sus apuntes. Si la informacion no esta, "
                 "dilo claramente. Cita la fuente con el nombre del archivo "
                 f"entre parentesis al final de cada idea.\n\n{contexto}"},
        {"text": "Pregunta: " + consulta},
    ])
    fuentes = "\n".join(f"• {rel(f)}" for _, f, _ in fragmentos)
    return f"{respuesta}\n\n— Notas consultadas —\n{fuentes}"


def ejecutar_accion(datos, confirmar, _interno=False):
    """Ejecuta una accion de la IA. Las que requieren permiso devuelven
    la propuesta y dejan el plan en PENDING si confirmar=False.
    _interno=True se usa desde el gestor de respuestas pendientes."""
    chat_id = datos.pop("_chat_id", None)
    action = datos.get("action", "reply_only")
    reply = datos.get("reply", "")

    if action == "reply_only":
        return reply or "(sin respuesta)"

    if action == "ask_folder":
        if not confirmar and chat_id is not None:
            PENDING[chat_id] = {"kind": "folder_choice", "data": datos}
            opciones = datos.get("options", [])
            lineas = [datos.get("reply", "No tengo claro dónde guardarlo."), ""]
            for i, carpeta in enumerate(opciones, 1):
                lineas.append(f"{i}. {carpeta if carpeta else '(raíz de la bóveda)'}")
            lineas.append("\nResponde con el número o el nombre de la carpeta.")
            return "\n".join(lineas)
        return "Pendiente de elegir carpeta."

    # --- Acciones directas (sin confirmacion) ---
    if action == "create_note":
        destino = safe_path(datos.get("folder", ""), datos.get("filename", ""))
        if destino.exists() and not confirmar and not _interno and chat_id is not None:
            PENDING[chat_id] = {"kind": "collision", "data": datos}
            return (f"🛑 Ya existe '{rel(destino)}'. ¿Qué hago?\n"
                    "1. Añadir el contenido como sección nueva con fecha\n"
                    "2. Crear otro archivo con nombre distinto\n"
                    "3. Cancelar\n\nResponde 1, 2 o 3.")
        if destino.exists():
            # Llegamos aqui tras confirmar colision tipo "seccion"
            with open(destino, "a", encoding="utf-8") as f:
                f.write(f"\n\n## Añadido el {datetime.now():%Y-%m-%d %H:%M}\n\n")
                f.write(datos.get("content", "").strip() + "\n")
            aviso = "(añadido como sección nueva)"
        else:
            destino.parent.mkdir(parents=True, exist_ok=True)
            destino.write_text(datos.get("content", "").strip() + "\n",
                               encoding="utf-8")
            aviso = ""
        ok, info = git_guardar(f"bot: crea {rel(destino)}")
        return f"{reply}\n\nCarpeta: {rel(destino)} {aviso}\n{info}"

    if action == "append_note":
        destino = find_note(datos.get("target_note", ""))
        if not destino:
            return f"No encuentro la nota '{datos.get('target_note')}'."
        with open(destino, "a", encoding="utf-8") as f:
            f.write(f"\n\n## Añadido el {datetime.now():%Y-%m-%d %H:%M}\n\n")
            f.write(datos.get("content", "").strip() + "\n")
        ok, info = git_guardar(f"bot: anade a {rel(destino)}")
        return f"{reply}\n\nAñadido a: {rel(destino)}\n{info}"

    # --- Acciones que requieren confirmacion ---
    if not confirmar and chat_id is not None:
        PENDING[chat_id] = datos
        return _describir_propuesta(action, datos)
    if confirmar:
        return _aplicar_confirmada(datos)
    return _describir_propuesta(action, datos)


def _describir_propuesta(action, datos):
    if action == "propose_folder":
        return (f"🛑 Necesita carpeta nueva:\n{datos.get('folder')}\n\n"
                f"Motivo: {datos.get('reply','')}\n\n"
                "Responde 'confirmar' para crearla y guardar la nota, "
                "o 'cancelar'.")
    if action == "move_note":
        return (f"🛑 Quiere mover/renombrar:\n"
                f"{datos.get('source')}  →  {datos.get('destination')}\n\n"
                f"{datos.get('reply','')}\n\n"
                "Responde 'confirmar' o 'cancelar'.")
    if action == "edit_note":
        antes = datos.get("old_text", "")[:1500]
        despues = datos.get("new_text", "")[:1500]
        return ("🛑 Quiere editar una nota. ANTES:\n"
                f"---\n{antes}\n---\n\nDESPUES:\n---\n{despues}\n---\n\n"
                f"Nota: {datos.get('target_note')}\n"
                "Responde 'confirmar' o 'cancelar'.")
    if action == "organizar":
        lineas = ["Propuesta de organizacion (nada se mueve aun):\n"]
        for i, m in enumerate(datos.get("moves", []), 1):
            lineas.append(f"{i}. {m.get('from')} → {m.get('to')}\n"
                          f"   porque: {m.get('porque','')}")
        lineas.append("\nResponde 'confirmar' o 'cancelar'.")
        return "\n".join(lineas)
    return "Tienes una accion pendiente. 'confirmar' o 'cancelar'."


def _aplicar_confirmada(datos):
    action = datos.get("action")

    if action == "propose_folder":
        ruta = crear_carpeta(datos.get("folder", ""))
        destino = safe_path(datos.get("folder", ""), datos.get("filename", ""))
        destino.write_text(datos.get("content", "").strip() + "\n",
                           encoding="utf-8")
        ok, info = git_guardar(f"bot: carpeta y nota {ruta}")
        return f"Carpeta creada y nota guardada en {rel(destino)}\n{info}"

    if action == "move_note":
        origen, destino = mover_nota(datos.get("source", ""),
                                     datos.get("destination", ""))
        ok, info = git_guardar(f"bot: mueve {origen} -> {destino}")
        return f"Movido: {origen} → {destino}\n{info}"

    if action == "edit_note":
        nota = find_note(datos.get("target_note", ""))
        if not nota:
            return f"No encuentro la nota '{datos.get('target_note')}'."
        texto = nota.read_text(encoding="utf-8")
        old, new = datos.get("old_text", ""), datos.get("new_text", "")
        if old and old in texto:
            texto = texto.replace(old, new, 1)
        elif not old:
            texto = new  # reemplazo completo
        else:
            return ("El texto original ya no coincide (puede que la nota "
                    "cambiara). No he tocado nada.")
        nota.write_text(texto, encoding="utf-8")
        ok, info = git_guardar(f"bot: edita {rel(nota)}")
        return f"Nota editada: {rel(nota)}\n{info}"

    if action == "organizar":
        hechos, fallos = [], []
        for m in datos.get("moves", []):
            try:
                o, d = mover_nota(m.get("from", ""), m.get("to", ""))
                hechos.append(f"{o} → {d}")
            except Exception as e:
                fallos.append(f"{m.get('from')}: {e}")
        if not hechos:
            return "No se pudo hacer ningun movimiento:\n" + "\n".join(fallos)
        ok, info = git_guardar("bot: organizacion de boveda")
        cuerpo = "Movimientos hechos:\n" + "\n".join(f"• {h}" for h in hechos)
        if fallos:
            cuerpo += "\n\nPendientes:\n" + "\n".join(f"• {f}" for f in fallos)
        return f"{cuerpo}\n\n{info}"

    return "Accion desconocida, no he tocado nada."


def plan_organizar():
    datos = _parsear_json(gemini([
        {"text": f"Estructura actual:\n{vault_tree()}\n\nLas notas deben estar "
                 "en su asignatura (001- 4to ESO/FiQ|Historia|Ingles) y su "
                 "seccion (02- Temario, 03- Trabajos, 04- Ejercicios). "
                 "Devuelve JSON {\"moves\":[{\"from\":\"ruta\",\"to\":\"ruta "
                 "completa con archivo\",\"porque\":\"texto\"}]}, maximo 12, "
                 "solo notas claramente mal ubicadas; si esta todo bien, "
                 "{\"moves\":[]}."},
    ], json_mode=True))
    datos["action"] = "organizar"
    return datos


# ----------------------------------------------------------------------
# Comandos
# ----------------------------------------------------------------------
AYUDA = """Bot de tu boveda Obsidian.

Me puedes mandar:
• Texto o nota de voz: clasifico y guardo en su carpeta.
• Foto del cuaderno/pizarra: la leo, detecto asignatura y la archivo.
• PDF o Word: extraigo el contenido y lo convierto en nota.
• Enlace web o de YouTube: hago un resumen estructurado.

Comandos:
• /pregunta <duda> - respondo desde tus apuntes citando la fuente
• /organizar       - propongo mover notas mal ubicadas (confirmas tu)
• /test            - crea una nota de prueba sin IA (diagnostico)
• /ver <nota>      - muestro una nota
• /archivo <nota>  - envio el .md como documento
• /buscar <texto>  - busco coincidencias
• /ultimas         - notas modificadas recientes
• /carpetas        - arbol de la boveda
• /ayuda           - este mensaje

Crear carpetas, mover y editar siempre te pide confirmacion. Nunca borro nada.
"""


def handle_command(chat_id, text):
    cmd, _, arg = text.partition(" ")
    arg = arg.strip()

    if cmd in ("/start", "/ayuda", "/help"):
        send_message(chat_id, AYUDA)
    elif cmd == "/carpetas":
        send_message(chat_id, "Estructura actual:\n\n" + vault_tree())
    elif cmd == "/ultimas":
        notas = sorted(md_files(), key=lambda f: f.stat().st_mtime,
                       reverse=True)[:8]
        send_message(chat_id, "Ultimas notas:\n\n" +
                     "\n".join(f"• {rel(f)}" for f in notas))
    elif cmd == "/buscar":
        if not arg:
            send_message(chat_id, "Uso: /buscar palabra")
            return
        resultados = []
        for puntos, f, trozo in buscar_fragmentos(arg):
            resultados.append(f"{rel(f)}: {trozo}")
        send_message(chat_id, ("Resultados:\n\n" + "\n".join(resultados))
                     if resultados else "Sin coincidencias.")
    elif cmd == "/pregunta":
        if not arg:
            send_message(chat_id, "Uso: /pregunta tu duda")
            return
        send_message(chat_id, "Buscando en tus apuntes...")
        send_message(chat_id, responder_rag(arg))
    elif cmd == "/ver":
        nota = find_note(arg)
        if not nota:
            send_message(chat_id, f"No encuentro '{arg}'.")
            return
        send_message(chat_id, f"# {nota.stem}\n\n" +
                     nota.read_text(encoding="utf-8", errors="ignore"))
    elif cmd == "/archivo":
        nota = find_note(arg)
        if not nota:
            send_message(chat_id, f"No encuentro '{arg}'.")
            return
        send_document(chat_id, nota)
        send_message(chat_id, f"Archivo: {rel(nota)}")
    elif cmd == "/test":
        # Prueba la tuberia de archivos y Git sin depender de la IA.
        destino = safe_path("", "prueba-bot.md")
        destino.write_text(
            f"# Prueba del bot\n\nCreada el {datetime.now():%Y-%m-%d %H:%M}\n",
            encoding="utf-8")
        ok, info = git_guardar("bot: prueba de tuberia")
        send_message(chat_id,
                     f"Nota de prueba creada en la raiz: {rel(destino)}\n{info}")
    elif cmd == "/organizar":
        send_message(chat_id, "Revisando la boveda...")
        datos = plan_organizar()
        if not datos.get("moves"):
            send_message(chat_id, "No veo notas mal ubicadas. Todo en orden.")
            return
        PENDING[chat_id] = datos
        send_message(chat_id, _describir_propuesta("organizar", datos))
    else:
        send_message(chat_id, "Comando desconocido. Usa /ayuda.")


# ----------------------------------------------------------------------
# Enrutado de mensajes
# ----------------------------------------------------------------------
def _gestionar_pendiente(chat_id, text):
    if chat_id not in PENDING:
        return None
    bajo = text.lower().strip()
    pendiente = PENDING[chat_id]

    # Estructura nueva con "kind"
    if isinstance(pendiente, dict) and pendiente.get("kind") == "folder_choice":
        if bajo in ("cancelar", "no", "3") and len(pendiente["data"].get("options", [])) < 3:
            PENDING.pop(chat_id, None)
            return "Cancelado, no he guardado nada."
        datos = pendiente["data"]
        opciones = datos.get("options", [])
        carpeta = None
        if bajo.isdigit() and 1 <= int(bajo) <= len(opciones):
            carpeta = opciones[int(bajo) - 1]
        else:
            candidatas = [c for c in opciones
                          if bajo in c.lower() or (not c and "ra" in bajo)]
            carpeta = candidatas[0] if candidatas else None
        if carpeta is None:
            return ("No reconozco esa opción. Manda el número de la carpeta "
                    "o 'cancelar'.")
        PENDING.pop(chat_id, None)
        datos["action"] = "create_note"
        datos["folder"] = carpeta
        return ejecutar_accion(datos, confirmar=False, _interno=True)

    if isinstance(pendiente, dict) and pendiente.get("kind") == "collision":
        datos = pendiente["data"]
        if bajo in ("1", "seccion", "sección"):
            PENDING.pop(chat_id, None)
            return ejecutar_accion(datos, confirmar=True, _interno=True)
        if bajo in ("2", "nueva", "otro", "otra"):
            PENDING.pop(chat_id, None)
            destino = safe_path(datos.get("folder", ""), datos.get("filename", ""))
            base = destino.stem
            n = 2
            while True:
                candidato = destino.with_name(f"{base} ({n}).md")
                if not candidato.exists():
                    candidato.write_text(datos.get("content", "").strip() + "\n",
                                         encoding="utf-8")
                    ok, info = git_guardar(f"bot: crea {rel(candidato)}")
                    return f"Creada una nota nueva: {rel(candidato)}\n{info}"
                n += 1
        if bajo in ("3", "cancelar", "no"):
            PENDING.pop(chat_id, None)
            return "Cancelado. No he tocado nada."
        return "Responde 1 (sección), 2 (archivo nuevo) o 3 (cancelar)."

    # Estructura clasica: propuestas de carpeta/mover/editar/organizar
    if bajo in ("confirmar", "si", "ok", "sí"):
        datos = PENDING.pop(chat_id)
        return _aplicar_confirmada(datos)
    if bajo in ("cancelar", "no"):
        PENDING.pop(chat_id, None)
        return "Plan descartado. No he tocado nada."
    return ("Tienes una propuesta pendiente. Responde 'confirmar' o "
            "'cancelar' antes de enviar otra cosa.")


def main():
    if not BOT_TOKEN or not GEMINI_KEY or \
            "PON-AQUI" in BOT_TOKEN or "PON-AQUI" in GEMINI_KEY:
        print("Faltan BOT_TOKEN o GEMINI_API_KEY (variables de entorno o config.py).")
        sys.exit(1)
    if not VAULT.exists():
        print(f"No existe la boveda: {VAULT}")
        sys.exit(1)

    print(f"Bot v2 en marcha (max {MAX_RUNTIME_MINUTES or 'sin'} min). "
          "Ctrl+C para parar.")
    offset = None
    inicio = time.time()
    while True:
        if MAX_RUNTIME_MINUTES and (time.time() - inicio) / 60 >= MAX_RUNTIME_MINUTES:
            print("Tiempo maximo alcanzado: cierro para el relevo del siguiente turno.")
            break
        try:
            resp = requests.post(
                f"{TG_API}/getUpdates",
                json={"offset": offset, "timeout": 30},
                timeout=45,
            ).json()
        except requests.RequestException:
            time.sleep(5)
            continue

        if not resp.get("ok"):
            time.sleep(5)
            continue

        for upd in resp["result"]:
            offset = upd["update_id"] + 1
            msg = upd.get("message")
            if not msg:
                continue
            chat_id = msg["chat"]["id"]

            if ALLOWED_CHAT_ID == 0:
                print(f"[aviso] chat_id detectado: {chat_id} "
                      "(ponlo en config.py para que solo tu puedas usarlo)")
            elif chat_id != ALLOWED_CHAT_ID:
                send_message(chat_id, "Bot privado. Acceso denegado.")
                continue

            try:
                # Confirmaciones pendientes: tienen prioridad absoluta
                if "text" in msg:
                    pendiente = _gestionar_pendiente(chat_id, msg["text"].strip())
                    if pendiente is not None:
                        send_message(chat_id, pendiente)
                        continue

                if "text" in msg:
                    text = msg["text"].strip()
                    if text.startswith("/"):
                        if text.startswith("/start"):
                            send_message(chat_id,
                                         f"Tu chat ID es: {chat_id}\n\n{AYUDA}")
                        else:
                            handle_command(chat_id, text)
                        continue

                    # Enlaces web / YouTube dentro del mensaje
                    urls = URL_RE.findall(text)
                    if urls:
                        url = urls[0]
                        send_message(chat_id, "Leyendo el enlace...")
                        try:
                            if YOUTUBE_RE.search(url):
                                contenido = extraer_youtube(url)
                                contexto = f"Video de YouTube: {url}"
                            else:
                                titulo, contenido = extraer_web(url)
                                contexto = f"Pagina web: {titulo} ({url})"
                            send_message(chat_id, "Procesando...")
                            send_message(chat_id, procesar_texto(
                                f"{contexto}\n\n{contenido}",
                                contexto="El alumno manda este enlace; "
                                         "archiva un resumen bien estructurado.",
                                chat_id=chat_id, forzar_accion=True))
                        except Exception as e:
                            send_message(chat_id, f"No pude leer el enlace: {e}")
                        continue

                    send_message(chat_id, "Procesando...")
                    send_message(chat_id, procesar_texto(text, chat_id=chat_id))

                elif "voice" in msg:
                    send_message(chat_id, "Transcribiendo audio...")
                    audio = download_telegram_file(msg["voice"]["file_id"])
                    texto = transcribir_audio(audio)
                    send_message(chat_id,
                                 f'Entendido: "{texto}"\n\nGuardando...')
                    send_message(chat_id, procesar_texto(
                        texto,
                        contexto="Mensaje original dictado por voz.",
                        chat_id=chat_id, forzar_accion=True))

                elif "photo" in msg:
                    send_message(chat_id, "Leyendo la imagen...")
                    foto = msg["photo"][-1]  # mayor resolucion
                    img = download_telegram_file(foto["file_id"])
                    info = extraer_imagen(img, "image/jpeg")
                    # Guardar la foto en la carpeta de adjuntos
                    carpeta_adj = ruta_segura(ATTACHMENTS_DIR)
                    carpeta_adj.mkdir(parents=True, exist_ok=True)
                    nombre_img = f"foto-{datetime.now():%Y%m%d-%H%M%S}.jpg"
                    (carpeta_adj / nombre_img).write_bytes(img)
                    embed = f"![[{ATTACHMENTS_DIR}/{nombre_img}]]"
                    send_message(chat_id,
                                 f"Detectado: {info.get('tipo','?')} de "
                                 f"{info.get('asignatura_estimada','?')} "
                                 f"({info.get('tema_estimado','?')}). "
                                 "Archivando...")
                    respuesta = procesar_texto(
                        "Contenido de la foto:\n" + info.get("markdown", ""),
                        contexto="El alumno mando una FOTO. Datos detectados: "
                        + json.dumps({k: v for k, v in info.items()
                                      if k != "markdown"}, ensure_ascii=False),
                        chat_id=chat_id, forzar_accion=True, insertar=embed)
                    # Si queda una propuesta pendiente, inyecta tambien ahi
                    for p in list(PENDING.values()):
                        d = p.get("data") if isinstance(p, dict) else None
                        if d and "content" in d and embed not in d["content"]:
                            d["content"] = embed + "\n\n" + d["content"]
                    send_message(chat_id, respuesta)

                elif "document" in msg:
                    doc = msg["document"]
                    nombre = doc.get("file_name", "documento")
                    mime = doc.get("mime_type", "")
                    send_message(chat_id, f"Descargando {nombre}...")
                    raw = download_telegram_file(doc["file_id"])
                    if mime == "application/pdf" or nombre.lower().endswith(".pdf"):
                        info = extraer_pdf(raw)
                    elif nombre.lower().endswith(".docx"):
                        info = extraer_docx(raw, nombre)
                    elif nombre.lower().endswith(".doc"):
                        send_message(chat_id,
                                     "El formato .doc antiguo no se puede leer; "
                                     "exportalo a PDF o .docx y reenvialo.")
                        continue
                    else:
                        send_message(chat_id,
                                     "Solo proceso PDF y Word (.docx).")
                        continue
                    send_message(chat_id, "Contenido extraido. Archivando...")
                    send_message(chat_id, procesar_texto(
                        "Contenido del documento:\n" + info.get("markdown", ""),
                        contexto="El alumno mando un DOCUMENTO. Datos detectados: "
                        + json.dumps({k: v for k, v in info.items()
                                      if k != "markdown"}, ensure_ascii=False),
                        chat_id=chat_id, forzar_accion=True))

                else:
                    send_message(chat_id,
                                 "No se procesar ese tipo de mensaje. "
                                 "Valen texto, voz, fotos, PDF, Word y enlaces.")
            except Exception as e:
                send_message(chat_id, f"Algo ha fallado: {e}")
                print(f"[error] {e}")


if __name__ == "__main__":
    main()
