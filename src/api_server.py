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
    description="HTTP API for generating PowerPoint presentations and Word documents with cloud delivery.",
    version="0.1.0"
)

# ... [existing API endpoints omitted for brevity] ...

# NEW ENDPOINT: Word document generation and upload
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
