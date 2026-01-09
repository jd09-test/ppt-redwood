# PPT Generator Server – REST AI Slide Generation

This project is a modern, AI-friendly backend service for generating PowerPoint presentations (PPTX), AI-voiced narration (mp3), and Word documents (docx) using strict layout/template controls and rich JSON content. It is designed for integration with LLM platforms (such as Oracle Fusion AI Studio), no longer includes any MCP (Model Context Protocol) or FastMCP server logic, and is 100% RESTful and cloud deployable.

## Features

- **/generate_json_content** – Accepts user text/topic/theme, calls OpenAI or other LLM for slide content JSON generation (fully backend-compliant schema).
- **/generate_presentation** – Accepts structured JSON for slides, placeholder content, (and speaker_notes), renders strictly-formatted PPTX from your template, uploads to Cloudinary.
- **/generate_word_doc** – Creates a Word document (.docx) from supplied text, uploads to Cloudinary for download/sharing.
- **/generate_ppt_with_audio** – Accepts JSON presentation and gender for audio narration (“male”/“female”), synthesizes speaker_notes for each slide to separate mp3s via Edge TTS (configurable voice), uploads audio per slide to Cloudinary, and returns all audio links alongside the pptx.
- **Speaker notes ready** – Each slide can include speaker_notes (string, not a visible placeholder, but for narration/presenter).
- **Theme support** – Light/dark theme selection.
- **Customizable layouts** – Uses layouts_template.yaml to map all JSON layouts to your PowerPoint template slides and placeholders.
- **Production-ready RESTful API with FastAPI** – No MCP tooling or protocols required.

## Deployment

- Designed to run on Render, AWS, or any cloud/server:
  ```sh
  uvicorn src.api_server:app --host=0.0.0.0 --port=8000
  ```

- Edit/provide `pyproject.toml` for all Python dependencies (`fastapi`, `openai`, `edge-tts`, `cloudinary`, `python-pptx`, `docx`, etc.)

- Supply your own strict corporate PowerPoint template as `src/assets/Oracle_PPT-template_FY26_blank.pptx`.
- Store all layouts and placeholder schema in `src/assets/layouts_template.yaml`.

## Integration Workflow

1. **User/Fusion prompts**: User provides context, selects theme and optionally voice via Fusion AI Studio or upstream system.

2. **JSON slide generation**: LLM or script creates valid JSON with slides, placeholders, and (optionally) speaker_notes.

3. **POST JSON to backend**: The system POSTs this JSON to `/generate_presentation` or `/generate_ppt_with_audio` as needed.

4. **Backend renders PPTX, generates audio, uploads to Cloudinary**.

5. **Download/Share**: System returns PPTX and per-slide mp3 URLs to user/client for download, playback, or further processing.

## Example API request for per-slide audio:

```json
{
  "json_content": {
    "filename": "sample_audio_ppt.pptx",
    "theme_mode": "light",
    "audio": "male",
    "slides": [
      {
        "slide_number": 1,
        "layout": "Title_Pillar",
        "placeholders": [
          { "placeholder_name": "Title", "content": ["AI Voice Demo"] },
          { "placeholder_name": "Subhead", "content": ["Dynamic narration"] }
        ],
        "speaker_notes": "This is a test slide narration for slide 1."
      }
      // (add more slides, each with optional speaker_notes)
    ]
  }
}
```

## Cleaning Up Past MCP/Server Artifacts

- All FastMCP, MCP, @mcp tooling, and src/server.py have been removed.
- Manual/test/dev scripts referencing MCP tooling were pruned.
- This REST API project requires no MCP involvement and runs as a plain FastAPI cloud service.

## File Structure

- src/api_server.py: All FastAPI endpoints – PPT, JSON, Word, Audio.
- src/assets/layouts_template.yaml: Complete, backend-locked template description.
- src/assets/Oracle_PPT-template_FY26_blank.pptx: PowerPoint rendering template.
- src/model.py: Data models.
- src/utils.py: Layout parsing, mapping, formatting utilities.
- src/assets/prompt.txt: Production prompt for LLM/AI pipeline.
- src/dev/: Developer utilities and test scripts (e.g., print_layout_placeholders.py, embed_audio_per_slide.py).

## Additional Notes

- Audio embedding (icons) into PPTX is not officially supported by python-pptx; audio is hosted on Cloudinary and linked per-slide.
- For layout/placeholder debugging, use print_layout_placeholders.py or edit src/assets/layouts_template.yaml.

---
