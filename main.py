import os, json, time, logging, requests, textwrap
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
# openai import removed for free summarization

logger = logging.getLogger(__name__)

# ==============================================================
# 1️⃣  Búsqueda web (Google Custom Search)
# ==============================================================

def web_search(query, engine_id, api_key, limit=10):
    """Devuelve una lista de dicts con título, link y snippet."""
    url = "https://www.googleapis.com/customsearch/v1"
    params = dict(
        key=api_key,
        cx=engine_id,
        q=query
    )
    r = requests.get(url, params=params)
    r.raise_for_status()
    items = r.json().get("items", [])

    results = []
    for i, item in enumerate(items[:limit]):
        results.append({
            "rank": i + 1,
            "title": item.get("title", ""),
            "link": item.get("link", ""),
            "snippet": item.get("snippet", ""),
            "displayLink": item.get("displayLink", ""),
        })
    return results

# ==============================================================
# 2️⃣  Resumir con OpenAI (o DeepSeek)
# ==============================================================

def summarize_text(text, words=250):
    """Resumir texto usando Sumy (LSA)."""
    # Ensure NLTK punkt data
    try:
        import nltk
        nltk.data.find('tokenizers/punkt')
    except LookupError:
        nltk.download('punkt')
    # Parse text and summarize
    from sumy.parsers.plaintext import PlaintextParser
    from sumy.nlp.tokenizers import Tokenizer
    from sumy.summarizers.lsa import LsaSummarizer

    parser = PlaintextParser.from_string(text, Tokenizer("spanish"))
    summarizer = LsaSummarizer()
    # sumy expects sentences count. Estimate
    sent_count = max(1, int(words / 100))
    summary_sentences = summarizer(parser.document, sent_count)
    return " ".join(str(s) for s in summary_sentences)
    """Función de resumen. Cambia a la LLM que prefieras."""
    openai.api_key = os.getenv("OPENAI_API_KEY")
    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "user",
             "content": f"Resume este texto en aproximadamente {words} palabras: \n\n{text}"}
        ],
        temperature=0.4,
    )
    return response.choices[0].message.content.strip()

# ==============================================================
# 3️⃣  Generar presentación de PowerPoint
# ==============================================================

def create_pptx(sources, summary, config):
    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = config["presentation_title"]
    subtitle.text = f"Fecha: {time.strftime('%Y-%m-%d')}"

    # Slide 1 – Resumen general
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(5))
    tf = txBox.text_frame
    tf.text = "Resumen Ejecutivo"
    p = tf.add_paragraph()
    p.text = summary
    p.font.size = Pt(14)

    # Slides 2–n – Fuentes individuales
    for src in sources:
        slide = prs.slides.add_slide(slide_layout)
        tf = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(5))
        p = tf.text_frame
        p.text = f"{src['rank']}. {src['title']}"
        p.font.size = Pt(18)
        p.space_after = Pt(12)

        p = tf.add_paragraph()
        p.text = f"Link: {src['link']}"
        p.font.size = Pt(12)
        p.font.color.rgb = (255, 0, 0)

        p = tf.add_paragraph()
        p.text = src['snippet']
        p.font.size = Pt(12)

    out_path = os.path.join("market-research", "report.pptx")
    prs.save(out_path)
    return out_path

# ==============================================================
# 4️⃣  Orquestador principal
# ==============================================================

def run(**kwargs):
    """
    Espera por un argumento:
        query: texto a buscar (“reputaçión media digital”, etc.)
        action: 'search' (búsqueda y resumen) | 'report' (crear PPTx)
    """
    query = kwargs.get("query")
    action = kwargs.get("action", "search")
    if not query:
        return "ERROR: No fue proporcionado un query."
    
    # 1. Cargar configuraciones y secretos
    with open("config.json") as f: config = json.load(f)
    api_key = os.getenv("GOOGLE_API_KEY")
    engine_id = os.getenv("GOOGLE_CSE_ID")
    if not api_key or not engine_id:
        return "ERROR: Faltan credenciales de la API de Google."

    # 2. Búsqueda web
    results = web_search(query, engine_id, api_key, limit=config["query_limit"])
    if not results:
        return "No se encontraron fuentes relevantes."

    # 3. Generar resumen (premiar tanto tu tiempo como el LLM)
    combined_text = "\n\n".join([f"{r['title']}\n{r['snippet']}" for r in results])
    summary = summarize_text(combined_text, words=config["summary_words"])

    if action == "search":
        return {
            "query": query,
            "results": results,
            "summary": summary
        }
    elif action == "report":
        ppt_path = create_pptx(results[:config["max_sources"]], summary, config)
        return f"Informe generado: {ppt_path}"
    else:
        return "Acción no reconocida. Usa 'search' o 'report'."
