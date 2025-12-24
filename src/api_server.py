from dotenv import load_dotenv
load_dotenv()

from fastapi import FastAPI, UploadFile, File, HTTPException
from pydantic import BaseModel
from typing import List, Optional
import os
import json

from src.utils import (
    get_layouts,
    load_prompt_template,
    generate_timestamped_filename,
    ppt_template,
    process_html,
    get_reverse_index,
)
from src.model import PresentationContent
from pptx import Presentation

# NEW: OpenAI
import openai

# NEW: Cloudinary for file hosting
import cloudinary
import cloudinary.uploader

cloudinary.config(
    cloud_name=os.environ.get("CLOUDINARY_CLOUD_NAME"),
    api_key=os.environ.get("CLOUDINARY_API_KEY"),
    api_secret=os.environ.get("CLOUDINARY_API_SECRET"),
    secure=True
)

app = FastAPI(
    title="PPT Generator API",
    description="HTTP API for generating PowerPoint presentations following the MCP tool interface.",
    version="0.1.0"
)

@app.get("/get_presentation_rules")
async def get_presentation_rules():
    """Returns the rules prompt for generating a presentation, with layout definitions."""
    try:
        layouts_str = get_layouts()
        prompt_str = load_prompt_template()
        from string import Template
        prompt_template = Template(prompt_str)
        rule_prompt = prompt_template.substitute(layouts_description=layouts_str)
        return {"rule_prompt": rule_prompt}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating rules: {str(e)}")

class GeneratePresentationRequest(BaseModel):
    json_content: dict

@app.post("/generate_presentation")
async def generate_presentation(data: GeneratePresentationRequest):
    """Accepts JSON presentation content, generates PPTX, and returns download location."""
    try:
        prs_content = PresentationContent(**data.json_content)
        slides_content = prs_content.slides
        theme_mode = prs_content.theme_mode

        message = {"slides_count": 0}
        prs = Presentation(ppt_template)
        reverse_index = get_reverse_index()

        import re
        layout_keys = set(reverse_index.keys())
        def normalize_layout_key(in_key):
            if in_key in layout_keys:
                return in_key
            alt = in_key.replace("_", "/").strip()
            if alt in layout_keys:
                return alt
            for key in layout_keys:
                if key.replace(" ", "").replace("/", "").lower() == in_key.replace(" ", "").replace("_", "").lower():
                    return key
            return None

        for slide_content in slides_content:
            orig_layout = slide_content.layout
            layout_name = normalize_layout_key(orig_layout)
            if not layout_name:
                raise HTTPException(
                    status_code=422, 
                    detail=f"Layout '{orig_layout}' not recognized. Valid layouts: {sorted(layout_keys)}"
                )
            layout_index = reverse_index[layout_name]["layout_index"][theme_mode]
            placeholders_reverse = reverse_index[layout_name]["placeholders"]
            slide_layout = prs.slide_layouts[layout_index]
            slide = prs.slides.add_slide(slide_layout)
            for placeholder in slide_content.placeholders:
                name = placeholder.placeholder_name
                content = placeholder.content
                try:
                    index = placeholders_reverse[name][theme_mode]
                    ph = slide.placeholders.__getitem__(idx=index)
                    text_frame = ph.text_frame
                    text_frame.clear()
                    text_frame._element.remove(text_frame.paragraphs[0]._p)
                    if isinstance(content, list):
                        for c in content:
                            p = text_frame.add_paragraph()
                            process_html(c, p)
                except Exception as e:
                    break
            message["slides_count"] += 1

        # Save to in-memory buffer, upload to Cloudinary, do not store to disk
        import io
        pptx_buffer = io.BytesIO()
        prs.save(pptx_buffer)
        pptx_buffer.seek(0)

        try:
            upload_result = cloudinary.uploader.upload(
                pptx_buffer,
                resource_type="raw",
                public_id=prs_content.filename,
                folder="presentations"
            )
            message["cloudinary_url"] = upload_result.get("secure_url")
            message["cloudinary_public_id"] = upload_result.get("public_id")
        except Exception as cloud_err:
            message["cloudinary_url"] = None
            message["cloudinary_error"] = str(cloud_err)

        message["status"] = "success"
        message["presentation_json"] = data.json_content
        return message
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating PPTX: {str(e)}")

# NEW ENDPOINT: Generate JSON content using OpenAI
class GenerateJsonContentRequest(BaseModel):
    user_text: str
    theme_mode: Optional[str] = "light"  # "light" or "dark"
    # You could also add more options (e.g., filename, etc.)

@app.post("/generate_json_content")
async def generate_json_content(data: GenerateJsonContentRequest):
    """Call OpenAI to generate valid PresentationContent JSON for presentation creation."""
    openai.api_key = os.environ.get("OPENAI_API_KEY")
    if not openai.api_key:
        raise HTTPException(status_code=503, detail="OpenAI API key not set in environment variable OPENAI_API_KEY.")
    try:
        layouts_str = get_layouts()
        prompt_str = load_prompt_template()
        from string import Template
        prompt_template = Template(prompt_str)
        system_prompt = prompt_template.substitute(layouts_description=layouts_str)
        user_prompt = (
            f"Input for presentation: {data.user_text}\n"
            f"Please use theme_mode: '{data.theme_mode}'."
        )
        completion = openai.chat.completions.create(
            # Use a suitable model ID (e.g., gpt-3.5-turbo, gpt-4-turbo)
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.3,
            max_tokens=4096
        )
        ai_json = completion.choices[0].message.content
        # The model should return a JSON code block as instructed; extract and parse it
        import re
        # Remove code fencing if present
        pattern = r"```(?:json)?\s*([\s\S]+?)\s*```"
        match = re.search(pattern, ai_json)
        if match:
            json_text = match.group(1)
        else:
            json_text = ai_json
        result = json.loads(json_text)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"OpenAI completion error: {str(e)}")

# FULLY AUTOMATED ENDPOINT
@app.post("/auto_presentation")
async def auto_presentation(data: GenerateJsonContentRequest):
    """
    Given just context and theme, auto-generate presentation JSON via OpenAI, create PPTX, and return the results in one call.
    """
    # Step 1: Generate JSON content using OpenAI
    openai.api_key = os.environ.get("OPENAI_API_KEY")
    if not openai.api_key:
        raise HTTPException(status_code=503, detail="OpenAI API key not set in environment variable OPENAI_API_KEY.")

    try:
        layouts_str = get_layouts()
        prompt_str = load_prompt_template()
        from string import Template
        prompt_template = Template(prompt_str)
        system_prompt = prompt_template.substitute(layouts_description=layouts_str)
        user_prompt = (
            f"Input for presentation: {data.user_text}\n"
            f"Please use theme_mode: '{data.theme_mode}'."
        )
        completion = openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.3,
            max_tokens=4096
        )
        ai_json = completion.choices[0].message.content
        import re
        pattern = r"```(?:json)?\s*([\s\S]+?)\s*```"
        match = re.search(pattern, ai_json)
        if match:
            json_text = match.group(1)
        else:
            json_text = ai_json
        # Remove trailing commas in JSON objects/arrays and fix smart quotes (common GPT artifact)
        import re
        json_text_clean = re.sub(r',(\s*[\}\]])', r'\1', json_text)
        # Replace smart quotes and similar ASCII-unsafe chars with regular ones
        json_text_clean = json_text_clean.replace('\u201c', '"').replace('\u201d', '"').replace('\u2018', "'").replace('\u2019', "'")
        json_text_clean = json_text_clean.replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'")
        # Remove other common LLM artifacts
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"OpenAI completion error: {str(e)}")
    
    # Step 2: Create PPTX with the generated JSON
    try:
        prs_content = PresentationContent(**result_json)
        default_save_location = os.getenv('WORK_DIR', os.path.join(os.path.expanduser('~'), "Documents"))
        filename = generate_timestamped_filename(default_save_location, prs_content.filename, "pptx")
        message = {"filename": os.path.basename(filename), "save_location": default_save_location, "slides_count": 0}
        theme_mode = prs_content.theme_mode
        slides_content = prs_content.slides

        prs = Presentation(ppt_template)
        reverse_index = get_reverse_index()

        layout_keys = set(reverse_index.keys())
        def normalize_layout_key(in_key):
            if in_key in layout_keys:
                return in_key
            alt = in_key.replace("_", "/").strip()
            if alt in layout_keys:
                return alt
            for key in layout_keys:
                if key.replace(" ", "").replace("/", "").lower() == in_key.replace(" ", "").replace("_", "").lower():
                    return key
            return None

        for slide_content in slides_content:
            orig_layout = slide_content.layout
            layout_name = normalize_layout_key(orig_layout)
            if not layout_name:
                raise HTTPException(
                    status_code=422, 
                    detail=f"Layout '{orig_layout}' not recognized. Valid layouts: {sorted(layout_keys)}"
                )
            layout_index = reverse_index[layout_name]["layout_index"][theme_mode]
            placeholders_reverse = reverse_index[layout_name]["placeholders"]
            slide_layout = prs.slide_layouts[layout_index]
            slide = prs.slides.add_slide(slide_layout)
            for placeholder in slide_content.placeholders:
                name = placeholder.placeholder_name
                content = placeholder.content
                try:
                    index = placeholders_reverse[name][theme_mode]
                    ph = slide.placeholders.__getitem__(idx=index)
                    text_frame = ph.text_frame
                    text_frame.clear()
                    text_frame._element.remove(text_frame.paragraphs[0]._p)
                    if isinstance(content, list):
                        for c in content:
                            p = text_frame.add_paragraph()
                            process_html(c, p)
                except Exception as e:
                    break
            message["slides_count"] += 1

        import io
        pptx_buffer = io.BytesIO()
        prs.save(pptx_buffer)
        pptx_buffer.seek(0)
        message["status"] = "success"
        message["presentation_json"] = result_json

        # Upload to Cloudinary from in-memory buffer (no local disk write)
        try:
            upload_result = cloudinary.uploader.upload(
                pptx_buffer,
                resource_type="raw",  # since pptx is not an image
                public_id=os.path.splitext(os.path.basename(filename))[0],
                folder="presentations"
            )
            message["cloudinary_url"] = upload_result.get("secure_url")
            message["cloudinary_public_id"] = upload_result.get("public_id")
        except Exception as cloud_err:
            message["cloudinary_url"] = None
            message["cloudinary_error"] = str(cloud_err)

        # Remove local storage keys for a true cloud workflow
        message.pop("save_location", None)
        message.pop("filename", None)

        return message
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating PPTX: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("src.api_server:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), reload=True)
