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


def safe_filename(text: str) -> str:
    text = re.sub(r"[^a-zA-Z0-9_-]+", "_", text.strip())
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:80] or "corpus_artigos"


def command_text(cmd: list[str]) -> str:
    return " ".join(f'"{part}"' if " " in part else part for part in cmd)


def run_command(cmd: list[str], timeout_seconds: int, log_box) -> tuple[int, str]:
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
                log_box.code("".join(lines[-120:]), language="text")
            if process.poll() is not None:
                remainder = process.stdout.read()
                if remainder:
                    lines.append(remainder)
                    log_box.code("".join(lines[-120:]), language="text")
                return process.returncode or 0, "".join(lines)
            if time.monotonic() - started > timeout_seconds:
                process.terminate()
                lines.append(f"\n[tempo limite atingido: {timeout_seconds // 60} min]\n")
                log_box.code("".join(lines[-120:]), language="text")
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
    st.set_page_config(page_title="Robo de artigos academicos", page_icon=None, layout="wide")

    st.title("Robo de busca de artigos")
    st.caption("Busca artigos em fontes academicas, limpa abstracts e gera uma planilha Excel pronta para baixar.")

    with st.form("search_form"):
        keywords = st.text_area(
            "Palavras-chave ou consultas",
            value="digital retail neural network\nomnichannel retail graph neural network\nretail supply chain network",
            height=130,
        )

        c1, c2, c3, c4 = st.columns(4)
        per_query = c1.number_input("Artigos por consulta", min_value=1, max_value=100, value=25, step=1)
        workers = c2.number_input("Agentes paralelos", min_value=1, max_value=16, value=8, step=1)
        max_records = c3.number_input("Limite final", min_value=1, max_value=5000, value=500, step=50)
        timeout_minutes = c4.number_input("Tempo limite min.", min_value=1, max_value=120, value=20, step=1)

        c5, c6, c7 = st.columns([1, 1, 2])
        year_from_raw = c5.text_input("Ano inicial", value="")
        year_to_raw = c6.text_input("Ano final", value="")
        include_networks = c7.checkbox("Somar redes predefinidas", value=False)

        filename_base = st.text_input("Nome do arquivo", value="corpus_artigos_limpo")
        submitted = st.form_submit_button("Gerar Excel")

    if "excel_bytes" in st.session_state:
        st.download_button(
            "Baixar ultimo Excel gerado",
            data=st.session_state["excel_bytes"],
            file_name=st.session_state.get("excel_name", "corpus_artigos_limpo.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if not submitted:
        return

    if not keywords.strip():
        st.error("Digite pelo menos uma palavra-chave ou consulta.")
        return

    try:
        year_from = int(year_from_raw) if year_from_raw.strip() else None
        year_to = int(year_to_raw) if year_to_raw.strip() else None
    except ValueError:
        st.error("Ano inicial e ano final precisam ser numeros inteiros.")
        return

    safe_base = safe_filename(filename_base)
    timeout_seconds = int(timeout_minutes) * 60
    log_box = st.empty()

    with tempfile.TemporaryDirectory(prefix="robo_artigos_") as tmp:
        output_dir = Path(tmp)
        final_output, commands = build_commands(
            keywords=keywords,
            output_dir=output_dir,
            filename_base=safe_base,
            per_query=int(per_query),
            workers=int(workers),
            max_records=int(max_records),
            year_from=year_from,
            year_to=year_to,
            include_networks=include_networks,
        )

        all_logs: list[str] = []
        with st.status("Gerando planilha...", expanded=True) as status:
            for index, cmd in enumerate(commands, start=1):
                st.write(f"Etapa {index}/{len(commands)}")
                code, log = run_command(cmd, timeout_seconds, log_box)
                all_logs.append(log)
                if code != 0:
                    status.update(label="Execucao interrompida", state="error")
                    st.error("A geracao nao foi concluida. Veja o log acima.")
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

        st.success("Excel gerado com sucesso.")
        st.download_button(
            "Baixar Excel",
            data=excel_bytes,
            file_name=f"{safe_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
