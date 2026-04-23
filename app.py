"""
pip install gradio requests python-docx freeplane-io
"""
import tempfile
from pathlib import Path

import gradio as gr
from docx import Document as DocxDocument

from llm import DEFAULT_PROMPT, PRESET_PROMPTS, list_models, run_task
from outliner import process_files


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _preview(path: Path) -> str:
    if path.suffix == ".md":
        return path.read_text(encoding="utf-8")[:4000]
    doc = DocxDocument(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())[:4000]


def refresh_models():
    models = list_models()
    if not models:
        return gr.update(choices=["(Ollama unavailable)"], value="(Ollama unavailable)")
    return gr.update(choices=models, value=models[0])



def on_preset_change(preset):
    return PRESET_PROMPTS[preset]


# ---------------------------------------------------------------------------
# Core processing
# ---------------------------------------------------------------------------

def run(file, mode, bold, italic, underline, model, prompt, markdown_mode, single_prompt, want_docx, want_mm, progress=gr.Progress()):
    if not file:
        raise gr.Error("Upload at least one document.")
    if not want_docx and not want_mm:
        raise gr.Error("Select at least one output format.")

    files = file if isinstance(file, list) else [file]
    input_paths = [Path(f.name) for f in files]
    stem  = "_".join(p.stem for p in input_paths)[:80]
    tmpdir = Path(tempfile.mkdtemp())

    doc_ext = ".md" if (mode == "LLM Task" and markdown_mode) else ".docx"
    doc_out = tmpdir / f"{stem}_output{doc_ext}" if want_docx else None
    mm_out  = tmpdir / f"{stem}_output.mm"       if want_mm  else None

    if mode == "Extract Marks":
        if not bold and not italic and not underline:
            raise gr.Error("Select at least one mark type.")
        actions = {"select_bold": bold, "select_italic": italic, "select_underline": underline}
    else:  # LLM Task
        if not model or "unavailable" in model:
            raise gr.Error("No Ollama model available. Is Ollama running?")
        if "{document}" not in prompt:
            raise gr.Error("The prompt must contain the {document} placeholder.")
        actions = {"summarizer": lambda text: run_task(text, prompt, model),
                   "markdown_mode": markdown_mode,
                   "single_prompt": single_prompt}

    progress(0, desc="Processing document…")
    process_files(input_paths, actions, doc_out, mm_out)
    progress(1, desc="Done")

    preview = ""
    if doc_out and doc_out.exists():
        preview = _preview(doc_out)

    docx_value = str(doc_out) if doc_out and doc_out.exists() else None
    mm_value   = str(mm_out)  if mm_out  and mm_out.exists()  else None

    return docx_value, mm_value, preview or "(no preview)"


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

initial_models = list_models()
initial_model  = initial_models[0] if initial_models else "(Ollama unavailable)"

with gr.Blocks(title="Document Outliner", theme=gr.themes.Soft()) as app:
    gr.Markdown("# Document Outliner")

    file_input = gr.File(label="Input document(s) (.docx / .odt)", file_types=[".docx", ".odt"], file_count="multiple")

    mode_state = gr.State("LLM Task")

    with gr.Tabs():
        with gr.Tab("LLM Task") as tab_llm:
            with gr.Row():
                model_dd    = gr.Dropdown(
                    choices=initial_models, value=initial_model,
                    label="Ollama model", scale=4,
                )
                refresh_btn = gr.Button("↺ Refresh", scale=1, min_width=90)
            preset_dd  = gr.Dropdown(
                choices=list(PRESET_PROMPTS.keys()),
                value="Summarize",
                label="Preset prompt",
            )
            prompt_box = gr.Textbox(
                value=DEFAULT_PROMPT,
                label="Prompt  — use {document} as placeholder for the section text",
                lines=5,
            )
            markdown_cb  = gr.Checkbox(label="Send document as Markdown (preserves bold/italic)", value=True)
            single_prompt_cb = gr.Checkbox(label="Send full document in a single prompt (instead of section by section)", value=False)

        with gr.Tab("Extract Marks") as tab_marks:
            with gr.Row():
                bold_cb      = gr.Checkbox(label="Bold",      value=True)
                italic_cb    = gr.Checkbox(label="Italic",    value=False)
                underline_cb = gr.Checkbox(label="Underline", value=False)

    # -- Output format --
    gr.Markdown("**Output formats**")
    with gr.Row():
        want_docx = gr.Checkbox(label="Document",  value=True)
        want_mm   = gr.Checkbox(label="Outline",   value=True)

    run_btn = gr.Button("Process", variant="primary", size="lg")

    preview_box = gr.Textbox(label="Preview", lines=14, interactive=False)

    with gr.Row():
        out_docx = gr.File(label="Download Document")
        out_mm   = gr.File(label="Download Outline")

    # -- Event wiring --
    tab_llm.select(lambda: "LLM Task",       outputs=mode_state)
    tab_marks.select(lambda: "Extract Marks", outputs=mode_state)

    refresh_btn.click(refresh_models, outputs=model_dd)
    preset_dd.change(on_preset_change, inputs=preset_dd, outputs=prompt_box)

    run_btn.click(
        run,
        inputs=[file_input, mode_state, bold_cb, italic_cb, underline_cb,
                model_dd, prompt_box, markdown_cb, single_prompt_cb, want_docx, want_mm],
        outputs=[out_docx, out_mm, preview_box],
    )


if __name__ == "__main__":
    app.launch()
