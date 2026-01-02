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

import openai
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
    description="HTTP API for generating PowerPoint presentations and Word documents with cloud delivery.",
    version="0.1.0"
)

@app.get("/get_presentation_rules")
async def get_presentation_rules():
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
        return { "cloudinary_url": message["cloudinary_url"] }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating PPTX: {str(e)}")

# The rest of the endpoints (e.g. /generate_json_content, /auto_presentation) should remain unchanged

from docx import Document

class GenerateWordDocRequest(BaseModel):
    content: str
    filename: Optional[str] = "blog_post"

@app.post("/generate_word_doc")
async def generate_word_doc(data: GenerateWordDocRequest):
    """
    Accepts blog/article content (plain text), generates a .docx file, uploads to Cloudinary, returns download URL.
    """
    try:
        doc = Document()
        # For minimal MVP: split by double newlines for paragraphs, else use single block
        paragraphs = data.content.split("\n\n")
        for para in paragraphs:
            doc.add_paragraph(para.strip())

        # Save to in-memory buffer, upload to Cloudinary
        import io
        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)

        try:
            upload_result = cloudinary.uploader.upload(
                docx_buffer,
                resource_type="raw",
                public_id=data.filename,
                folder="word_docs"
            )
            cloud_url = upload_result.get("secure_url")
        except Exception as upload_err:
            raise HTTPException(status_code=500, detail=f"Cloudinary upload error: {str(upload_err)}")

        return { "cloudinary_url": cloud_url }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating DOCX: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("src.api_server:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), reload=True)
