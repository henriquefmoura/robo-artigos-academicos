from __future__ import annotations

import re
import subprocess
import sys
import tempfile
import time
import unicodedata
from datetime import datetime
from pathlib import Path

import requests
import streamlit as st


BASE = Path(__file__).resolve().parent
SEARCH_SCRIPT = BASE / "buscar_artigos_redes.py"
CLEAN_SCRIPT = BASE / "rpa_limpar_abstracts.py"
DEFAULT_PER_QUERY = 25
DEFAULT_WORKERS = 8
DEFAULT_MAX_RECORDS = 500
DEFAULT_TIMEOUT_MINUTES = 20
DEFAULT_INCLUDE_NETWORKS = False
TRANSLATION_GLOSSARY = {
    "aprendizado de maquina": "machine learning",
    "aprendizagem de maquina": "machine learning",
    "inteligencia artificial": "artificial intelligence",
    "redes neurais": "neural networks",
    "rede neural": "neural network",
    "redes neurais profundas": "deep neural networks",
    "aprendizado profundo": "deep learning",
    "aprendizagem profunda": "deep learning",
    "varejo": "retail",
    "varejo digital": "digital retail",
    "omnichannel": "omnichannel",
    "omnichannel varejo": "omnichannel retail",
    "cadeia de suprimentos": "supply chain",
    "logistica": "logistics",
    "comportamento do consumidor": "consumer behavior",
    "comportamento de compra": "shopping behavior",
    "experiencia do cliente": "customer experience",
    "satisfacao do cliente": "customer satisfaction",
    "fidelizacao de clientes": "customer loyalty",
    "transformacao digital": "digital transformation",
    "inovacao": "innovation",
    "sustentabilidade": "sustainability",
    "marketing digital": "digital marketing",
    "comercio eletronico": "e-commerce",
    "ecommerce": "e-commerce",
    "plataformas digitais": "digital platforms",
    "analise de redes": "network analysis",
    "redes sociais": "social networks",
    "grafos": "graphs",
    "redes bayesianas": "bayesian networks",
}


def safe_filename(text: str) -> str:
    text = re.sub(r"[^a-zA-Z0-9_-]+", "_", text.strip())
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:80] or f"planilha_artigos_{datetime.now():%Y%m%d}"


def normalize_lookup(text: str) -> str:
    text = unicodedata.normalize("NFKD", text.lower())
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip()


def translate_pt_to_en(text: str) -> tuple[str, str]:
    text = text.strip()
    if not text:
        return "", ""

    lookup = normalize_lookup(text)
    if lookup in TRANSLATION_GLOSSARY:
        return TRANSLATION_GLOSSARY[lookup], "glossario"

    try:
        response = requests.get(
            "https://api.mymemory.translated.net/get",
            params={"q": text, "langpair": "pt|en"},
            timeout=8,
        )
        if response.status_code == 200:
            data = response.json()
            translated = (data.get("responseData") or {}).get("translatedText", "")
            translated = re.sub(r"\s+", " ", translated).strip()
            if translated:
                return translated, "tradutor online"
    except Exception:
        pass

    words = [TRANSLATION_GLOSSARY.get(normalize_lookup(part), part) for part in re.split(r"[,;]", text)]
    translated = ", ".join(part.strip() for part in words if part.strip())
    return translated or text, "texto original"


def command_text(cmd: list[str]) -> str:
    return " ".join(f'"{part}"' if " " in part else part for part in cmd)


def run_command(cmd: list[str], timeout_seconds: int) -> tuple[int, str]:
    started = time.monotonic()
    lines: list[str] = [f"$ {command_text(cmd)}\n"]
    process = subprocess.Popen(
        cmd,
        cwd=str(BASE),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    assert process.stdout is not None

    try:
        while True:
            line = process.stdout.readline()
            if line:
                lines.append(line)
            if process.poll() is not None:
                remainder = process.stdout.read()
                if remainder:
                    lines.append(remainder)
                return process.returncode or 0, "".join(lines)
            if time.monotonic() - started > timeout_seconds:
                process.terminate()
                lines.append(f"\n[tempo limite atingido: {timeout_seconds // 60} min]\n")
                return 124, "".join(lines)
            time.sleep(0.1)
    finally:
        if process.poll() is None:
            process.kill()


def build_commands(
    keywords: str,
    output_dir: Path,
    filename_base: str,
    per_query: int,
    workers: int,
    max_records: int,
    year_from: int | None,
    year_to: int | None,
    include_networks: bool,
) -> tuple[Path, list[list[str]]]:
    raw_output = output_dir / f"{filename_base}_bruto.xlsx"
    final_output = output_dir / f"{filename_base}.xlsx"

    search_cmd = [
        sys.executable,
        "-u",
        str(SEARCH_SCRIPT),
        "--mode",
        "both" if include_networks else "keywords",
        "--keywords",
        keywords,
        "--per-query",
        str(per_query),
        "--workers",
        str(workers),
        "--max-records",
        str(max_records),
        "--wait",
        "0.05",
        "--output",
        str(raw_output),
    ]
    if year_from:
        search_cmd.extend(["--year-from", str(year_from)])
    if year_to:
        search_cmd.extend(["--year-to", str(year_to)])

    clean_cmd = [
        sys.executable,
        "-u",
        str(CLEAN_SCRIPT),
        "--input",
        str(raw_output),
        "--output",
        str(final_output),
        "--wait",
        "0.05",
        "--no-auto-reference",
        "--drop-empty-abstracts",
        "--skip-missing-fetch",
    ]
    return final_output, [search_cmd, clean_cmd]


def apply_theme() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background: linear-gradient(180deg, #f6fbff 0%, #ffffff 46%, #f8fafc 100%);
        }
        .block-container {
            padding-top: 2rem;
            padding-bottom: 3rem;
            max-width: 980px;
        }
        .hero {
            border-radius: 18px;
            padding: 30px 34px;
            margin-bottom: 24px;
            background:
                linear-gradient(135deg, rgba(13, 148, 136, 0.94), rgba(37, 99, 235, 0.92) 58%, rgba(245, 158, 11, 0.94));
            color: white;
            box-shadow: 0 18px 40px rgba(15, 23, 42, 0.16);
        }
        .hero h1 {
            color: white;
            font-size: 2.2rem;
            line-height: 1.1;
            margin: 0 0 10px;
            letter-spacing: 0;
        }
        .hero p {
            color: rgba(255, 255, 255, 0.94);
            font-size: 1.02rem;
            margin: 0;
            max-width: 760px;
        }
        .eyebrow {
            display: inline-block;
            padding: 5px 10px;
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.18);
            color: white;
            font-weight: 700;
            font-size: 0.78rem;
            margin-bottom: 14px;
        }
        .step-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 12px;
            margin: 0 0 22px;
        }
        .step-card {
            border: 1px solid #dbeafe;
            border-radius: 12px;
            padding: 14px 16px;
            background: white;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
        }
        .step-number {
            color: #0f766e;
            font-weight: 800;
            font-size: 0.82rem;
            margin-bottom: 4px;
        }
        .step-card strong {
            display: block;
            color: #0f172a;
            margin-bottom: 3px;
        }
        .step-card span {
            color: #475569;
            font-size: 0.9rem;
        }
        .section-copy {
            color: #475569;
            margin: -6px 0 14px;
        }
        .stTextInput input {
            border-radius: 10px;
            border: 1px solid #cbd5e1;
            background: #ffffff;
        }
        .stTextInput input:focus {
            border-color: #0d9488;
            box-shadow: 0 0 0 1px #0d9488;
        }
        .stButton button, .stDownloadButton button {
            border-radius: 999px;
            border: 0;
            background: linear-gradient(90deg, #0f766e, #2563eb);
            color: white;
            font-weight: 800;
            padding: 0.7rem 1.25rem;
        }
        .stButton button:hover, .stDownloadButton button:hover {
            color: white;
            filter: brightness(1.05);
        }
        [data-testid="stFormSubmitButton"] button {
            width: 100%;
            min-height: 56px;
            margin-top: 10px;
            font-size: 1.05rem;
            letter-spacing: 0;
            background: linear-gradient(90deg, #f59e0b, #0d9488 45%, #2563eb);
            box-shadow: 0 14px 30px rgba(37, 99, 235, 0.22);
        }
        [data-testid="stFormSubmitButton"] button:hover {
            box-shadow: 0 16px 36px rgba(13, 148, 136, 0.28);
            transform: translateY(-1px);
        }
        [data-testid="stForm"] {
            border: 1px solid #dbeafe;
            border-radius: 16px;
            padding: 22px;
            background: rgba(255, 255, 255, 0.92);
            box-shadow: 0 10px 30px rgba(15, 23, 42, 0.06);
        }
        @media (max-width: 760px) {
            .step-grid {
                grid-template-columns: 1fr;
            }
            .hero {
                padding: 24px;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(page_title="Robo de artigos academicos", page_icon=None, layout="centered")
    apply_theme()
    for index in range(1, 6):
        st.session_state.setdefault(f"keyword_{index}", "")

    st.markdown(
        """
        <div class="hero">
            <div class="eyebrow">Busca academica automatizada</div>
            <h1>Gerador de planilha de artigos academicos</h1>
            <p>Informe seus termos de pesquisa. O aplicativo busca artigos em bases abertas, limpa os abstracts e entrega um Excel padronizado para baixar.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="step-grid">
            <div class="step-card"><div class="step-number">Etapa 1</div><strong>Busca</strong><span>Consulta bases academicas abertas.</span></div>
            <div class="step-card"><div class="step-number">Etapa 2</div><strong>Organiza</strong><span>Remove duplicados e abstracts invalidos.</span></div>
            <div class="step-card"><div class="step-number">Etapa 3</div><strong>Entrega</strong><span>Gera uma planilha Excel pronta.</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    input_col, translator_col = st.columns([1.55, 1], gap="large")

    with input_col:
        with st.form("search_form"):
            st.subheader("Tema ou palavras-chave")
            st.markdown('<p class="section-copy">Insira ate 5 palavras-chave ou consultas. Cada campo preenchido gera uma busca separada.</p>', unsafe_allow_html=True)
            for index in range(1, 6):
                st.text_input(
                    f"Palavra-chave {index}",
                    key=f"keyword_{index}",
                    placeholder="Ex.: artificial intelligence in education",
                )

            c1, c2 = st.columns(2)
            year_from_raw = c1.text_input("Ano inicial opcional", value="", placeholder="Ex.: 2020")
            year_to_raw = c2.text_input("Ano final opcional", value="", placeholder="Ex.: 2026")

            filename_base = st.text_input(
                "Nome do arquivo",
                value="planilha_artigos_academicos",
                placeholder="Ex.: revisao_inteligencia_artificial",
                help="Voce pode manter o nome sugerido ou escrever outro. A extensao .xlsx sera adicionada automaticamente.",
            )
            st.info(
                "Ao clicar, o aplicativo vai consultar fontes academicas abertas, combinar resultados repetidos, "
                "limpar abstracts e montar uma planilha Excel padronizada para download."
            )
            submitted = st.form_submit_button("Gerar minha planilha Excel pronta")

    with translator_col:
        with st.container(border=True):
            st.subheader("Tradutor rapido")
            st.caption("Digite uma palavra ou expressao em portugues para usar na busca em ingles.")
            pt_text = st.text_input("Termo em portugues", placeholder="Ex.: varejo digital")
            if st.button("Traduzir para ingles", use_container_width=True):
                translated, source = translate_pt_to_en(pt_text)
                st.session_state["translated_keyword"] = translated
                st.session_state["translated_source"] = source

            translated_keyword = st.session_state.get("translated_keyword", "")
            if translated_keyword:
                st.success(translated_keyword)
                st.caption(f"Fonte: {st.session_state.get('translated_source', 'tradutor')}")
                if st.button("Inserir no proximo campo livre", use_container_width=True):
                    inserted = False
                    for index in range(1, 6):
                        key = f"keyword_{index}"
                        if not st.session_state.get(key, "").strip():
                            st.session_state[key] = translated_keyword
                            inserted = True
                            break
                    if not inserted:
                        st.session_state["keyword_5"] = translated_keyword
                    st.rerun()

    if "excel_bytes" in st.session_state:
        st.download_button(
            "Baixar ultimo Excel gerado",
            data=st.session_state["excel_bytes"],
            file_name=st.session_state.get("excel_name", "planilha_artigos_academicos.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if not submitted:
        return

    keywords = "\n".join(st.session_state.get(f"keyword_{index}", "").strip() for index in range(1, 6) if st.session_state.get(f"keyword_{index}", "").strip())
    if not keywords.strip():
        st.error("Digite pelo menos um tema ou palavra-chave.")
        return

    try:
        year_from = int(year_from_raw) if year_from_raw.strip() else None
        year_to = int(year_to_raw) if year_to_raw.strip() else None
    except ValueError:
        st.error("Ano inicial e ano final precisam ser numeros inteiros.")
        return

    safe_base = safe_filename(filename_base)
    timeout_seconds = DEFAULT_TIMEOUT_MINUTES * 60

    with tempfile.TemporaryDirectory(prefix="robo_artigos_") as tmp:
        output_dir = Path(tmp)
        final_output, commands = build_commands(
            keywords=keywords,
            output_dir=output_dir,
            filename_base=safe_base,
            per_query=DEFAULT_PER_QUERY,
            workers=DEFAULT_WORKERS,
            max_records=DEFAULT_MAX_RECORDS,
            year_from=year_from,
            year_to=year_to,
            include_networks=DEFAULT_INCLUDE_NETWORKS,
        )

        all_logs: list[str] = []
        with st.status("Preparando sua planilha...", expanded=False) as status:
            for index, cmd in enumerate(commands, start=1):
                if index == 1:
                    st.write("Buscando artigos e abstracts nas bases academicas.")
                else:
                    st.write("Limpando e organizando a planilha final.")
                code, log = run_command(cmd, timeout_seconds)
                all_logs.append(log)
                if code != 0:
                    status.update(label="Execucao interrompida", state="error")
                    st.error("A planilha nao foi concluida. Tente uma busca mais especifica ou execute novamente.")
                    st.session_state["last_log"] = "\n".join(all_logs)
                    return
            status.update(label="Planilha pronta", state="complete")

        if not final_output.exists():
            st.error("O arquivo final nao foi encontrado apos a execucao.")
            return

        excel_bytes = final_output.read_bytes()
        st.session_state["excel_bytes"] = excel_bytes
        st.session_state["excel_name"] = f"{safe_base}.xlsx"
        st.session_state["last_log"] = "\n".join(all_logs)

        st.success("Sua planilha esta pronta.")
        st.download_button(
            "Baixar planilha Excel",
            data=excel_bytes,
            file_name=f"{safe_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
