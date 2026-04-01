This application reads a `docx` word document and outputs a new document that preserves document structure (headings) but replaces content with:

1. An **LLM-generated result** for each section — summary, key points, translation, or any task defined via a custom prompt — powered by a local [Ollama](https://ollama.com/) instance. OR
2. A **compilation** of bold/underlined segments (annotated by the user to highlight the most relevant terms or ideas)

Output is available in two formats (one or both at the same time):

- a new document (Word `.docx`)
- a concept map ([Freeplane](https://docs.freeplane.org/) `.mm`)

This project has been highly motivated by learning activities in order to avoid a second reading.

## Setup

### Virtual environment

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
```

### Dependencies

```bash
pip install gradio requests python-docx freeplane-io
```

### Ollama

Requires a running Ollama instance. The recommended setup with Docker (enables browser CORS):

```bash
docker run -d --name ollama_API --restart always \
  -v ~/.IAmodels:/root/.ollama -p 11434:11434 --gpus all \
  -e OLLAMA_ORIGINS='*' ollama/ollama
```

Pull a model before using the app, e.g.:

```bash
docker exec ollama_API ollama pull llama3.2
```

## Web interface

```bash
python app.py
```

Open `http://localhost:7860` in the browser.

![Gradio UI](./capture.png)

### Modes

**LLM Task** — select an Ollama model, choose a preset prompt or write your own using `{document}` as the placeholder for each section's text:

| Preset | Effect |
|---|---|
| Summarize | Concise summary of each section |
| Key points | Bullet list of main ideas |
| Critical analysis | Strengths and weaknesses |
| Translate to Spanish / English | Section-by-section translation |

**Extract Marks** — collects bold and/or underlined runs from each paragraph.

Both output formats (`.docx` and `.mm`) can be downloaded simultaneously after processing.

## Samples

### Original excerpt

![original](./img/original_excerpt.png)

### LLM Task (summarization)

![freeplane_summary](./img/freeplane_summary.png)

![word_summary](./img/word_summary.png)

### Marked words (bold, underlined)

![freeplane_marked](./img/freeplane_marked.png)

![word_marked](./img/word_marked.png)

## Command line usage

```bash
python outliner.py <input> [options]
```

| Option | Description |
|---|---|
| `-b` / `--bold` | Extract bold text |
| `-u` / `--underline` | Extract underlined text |
| `-s` / `--summary` | Generate LLM output per section |
| `-m` / `--model` | Ollama model name (default: `llama3.2`) |
| `-d` / `--docx <path>` | Word output path (file or folder) |
| `-f` / `--freeplane <path>` | Freeplane output path (file or folder) |
| `-r` / `--recursive` | Process subfolders (folder input only) |

### Examples

Extract bold and underlined into a Freeplane file:

```bash
python outliner.py ./data/sample.docx -b -u --freeplane ./data/output
```

Generate an LLM summary into both Word and Freeplane:

```bash
python outliner.py ./data/sample.docx -s --model llama3.2 --docx ./data/output --freeplane ./data/output
```

Batch process a folder recursively:

```bash
python outliner.py ./data -r -b -u --docx ./data/output --freeplane ./data/output
```

## TO-DO

- [x] Summarization
- [x] Local LLM via Ollama (any task via custom prompt)
- [x] Web interface (Gradio)
- [ ] Highlighted text (yellow, …)
- [ ] ODF text documents (`.odt`)
- [ ] Use topic models
- [ ] Text to speech for more passive readings
