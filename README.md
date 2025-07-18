# LLM Word Agent

A lightweight, LLM-powered Word document editing toolkit for Python.

## Features

- **docx_utils.py**: Full-featured docx utility
  - Create, open, save Word documents
  - Paragraph, heading, page break, list, table, image, hyperlink, header/footer, section, style
  - Batch replace, merge, split, extract, export structure for LLM/agent
- **llm_agent.py**: LLM-driven Word document agent
  - Automatically generates and executes editing commands based on user instructions
  - Supports OpenAI-compatible LLMs (OpenAI, Azure, OpenRouter, etc)
  - Demo and usage included

## Installation

```bash
git clone <your-github-repo-url>
cd <project-folder>
conda create -n llmword python=3.9
conda activate llmword
pip install -r requirements.txt
```

## Usage

- See `llm_agent.py` for main usage and examples.
- You can run the script directly for a demo (edit the API key and file paths as needed):

```bash
python llm_agent.py
```

## File Descriptions

- `docx_utils.py`: Full-featured docx operation toolkit.
- `llm_agent.py`: LLM-driven Word document agent, with demo and usage.

## Requirements

- python-docx
- openai

## License

MIT 