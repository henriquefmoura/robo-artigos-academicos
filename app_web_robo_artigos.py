from __future__ import annotations

import re
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import streamlit as st


BASE = Path(__file__).resolve().parent
SEARCH_SCRIPT = BASE / "buscar_artigos_redes.py"
CLEAN_SCRIPT = BASE / "rpa_limpar_abstracts.py"
DEFAULT_PER_QUERY = 25
DEFAULT_WORKERS = 8
DEFAULT_MAX_RECORDS = 500
DEFAULT_TIMEOUT_MINUTES = 20
DEFAULT_INCLUDE_NETWORKS = False


def safe_filename(text: str) -> str:
    text = re.sub(r"[^a-zA-Z0-9_-]+", "_", text.strip())
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:80] or "corpus_artigos"


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


def main() -> None:
    st.set_page_config(page_title="Robo de artigos academicos", page_icon=None, layout="centered")

    st.title("Gerador de planilha de artigos academicos")
    st.write(
        "Informe o tema da sua pesquisa. O aplicativo busca artigos em bases academicas abertas, "
        "remove registros sem abstract valido, organiza os metadados e entrega uma planilha Excel pronta para uso."
    )

    with st.form("search_form"):
        st.subheader("Tema ou palavras-chave")
        st.caption("Insira ate 5 temas/consultas. Cada campo gera uma busca separada.")
        keyword_defaults = [
            "digital retail neural network",
            "omnichannel retail graph neural network",
            "retail supply chain network",
            "",
            "",
        ]
        keyword_values: list[str] = []
        for index, default in enumerate(keyword_defaults, start=1):
            keyword_values.append(
                st.text_input(
                    f"Busca {index}",
                    value=default,
                    placeholder="Ex.: artificial intelligence in education",
                )
            )

        c1, c2 = st.columns(2)
        year_from_raw = c1.text_input("Ano inicial opcional", value="", placeholder="Ex.: 2020")
        year_to_raw = c2.text_input("Ano final opcional", value="", placeholder="Ex.: 2026")

        filename_base = st.text_input("Nome do arquivo", value="corpus_artigos_limpo")
        st.info(
            "Ao clicar, o aplicativo vai consultar fontes academicas abertas, combinar resultados repetidos, "
            "limpar abstracts e montar uma planilha Excel padronizada para download."
        )
        submitted = st.form_submit_button("Gerar planilha Excel pronta")

    if "excel_bytes" in st.session_state:
        st.download_button(
            "Baixar ultimo Excel gerado",
            data=st.session_state["excel_bytes"],
            file_name=st.session_state.get("excel_name", "corpus_artigos_limpo.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if not submitted:
        return

    keywords = "\n".join(value.strip() for value in keyword_values if value.strip())
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
