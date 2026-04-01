"""
pip install gradio requests python-docx freeplane-io
"""
import tempfile
from pathlib import Path

import gradio as gr
from docx import Document as DocxDocument

from llm import DEFAULT_PROMPT, PRESET_PROMPTS, list_models, run_task
from outliner import process_file


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _preview_docx(path: Path) -> str:
    doc = DocxDocument(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())[:4000]


def refresh_models():
    models = list_models()
    if not models:
        return gr.update(choices=["(Ollama unavailable)"], value="(Ollama unavailable)")
    return gr.update(choices=models, value=models[0])


def on_mode_change(mode):
    is_llm = mode == "LLM Task"
    return gr.update(visible=not is_llm), gr.update(visible=is_llm)


def on_preset_change(preset):
    return PRESET_PROMPTS[preset]


# ---------------------------------------------------------------------------
# Core processing
# ---------------------------------------------------------------------------

def run(file, mode, bold, underline, model, prompt, want_docx, want_mm, progress=gr.Progress()):
    if file is None:
        raise gr.Error("Upload a .docx file first.")
    if not want_docx and not want_mm:
        raise gr.Error("Select at least one output format.")

    input_path = Path(file.name)
    tmpdir     = Path(tempfile.mkdtemp())
    stem       = input_path.stem

    doc_out = tmpdir / f"{stem}_output.docx" if want_docx else None
    mm_out  = tmpdir / f"{stem}_output.mm"   if want_mm  else None

    if mode == "Extract Marks":
        if not bold and not underline:
            raise gr.Error("Select at least Bold or Underline.")
        actions = {"select_bold": bold, "select_underline": underline}
    else:  # LLM Task
        if not model or "unavailable" in model:
            raise gr.Error("No Ollama model available. Is Ollama running?")
        if "{document}" not in prompt:
            raise gr.Error("The prompt must contain the {document} placeholder.")
        actions = {"summarizer": lambda text: run_task(text, prompt, model)}

    progress(0, desc="Processing document…")
    process_file(input_path, actions, doc_out, mm_out)
    progress(1, desc="Done")

    preview = ""
    if doc_out and doc_out.exists():
        preview = _preview_docx(doc_out)

    docx_update = gr.update(visible=bool(doc_out and doc_out.exists()),
                            value=str(doc_out) if doc_out and doc_out.exists() else None)
    mm_update   = gr.update(visible=bool(mm_out  and mm_out.exists()),
                            value=str(mm_out)  if mm_out  and mm_out.exists()  else None)

    return docx_update, mm_update, preview or "(no preview)"


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

initial_models = list_models()
initial_model  = initial_models[0] if initial_models else "(Ollama unavailable)"

with gr.Blocks(title="Document Outliner", theme=gr.themes.Soft()) as app:
    gr.Markdown("# Document Outliner")

    file_input = gr.File(label="Input document (.docx)", file_types=[".docx"])

    mode = gr.Radio(
        ["Extract Marks", "LLM Task"],
        value="LLM Task",
        label="Mode",
    )

    # -- Marks options (hidden by default) --
    with gr.Group(visible=False) as marks_group:
        with gr.Row():
            bold_cb      = gr.Checkbox(label="Bold",      value=True)
            underline_cb = gr.Checkbox(label="Underline", value=False)

    # -- LLM options --
    with gr.Group(visible=True) as llm_group:
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

    # -- Output format --
    gr.Markdown("**Output formats**")
    with gr.Row():
        want_docx = gr.Checkbox(label="Word (.docx)",    value=True)
        want_mm   = gr.Checkbox(label="Freeplane (.mm)", value=True)

    run_btn = gr.Button("Process", variant="primary", size="lg")

    preview_box = gr.Textbox(label="Preview", lines=14, interactive=False)

    with gr.Row():
        out_docx = gr.File(label="Download .docx", visible=False)
        out_mm   = gr.File(label="Download .mm",   visible=False)

    # -- Event wiring --
    mode.change(on_mode_change, inputs=mode, outputs=[marks_group, llm_group])
    refresh_btn.click(refresh_models, outputs=model_dd)
    preset_dd.change(on_preset_change, inputs=preset_dd, outputs=prompt_box)

    run_btn.click(
        run,
        inputs=[file_input, mode, bold_cb, underline_cb,
                model_dd, prompt_box, want_docx, want_mm],
        outputs=[out_docx, out_mm, preview_box],
    )


if __name__ == "__main__":
    app.launch()
