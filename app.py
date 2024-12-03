from flask import Flask, request, jsonify, send_from_directory
from flask import render_template
from pathlib import Path
import json
import yaml
import os
from datetime import datetime

from docling_core.types.doc import ImageRefMode, PictureItem, TableItem
from docling.backend.pypdfium2_backend import PyPdfiumDocumentBackend
from docling.backend.asciidoc_backend import AsciiDocBackend
from docling.backend.msexcel_backend import MsExcelDocumentBackend
from docling.backend.msword_backend import MsWordDocumentBackend
from docling.backend.mspowerpoint_backend import MsPowerpointDocumentBackend
from docling.backend.md_backend import MarkdownDocumentBackend
from docling.backend.html_backend import HTMLDocumentBackend
from docling.backend.docling_parse_backend import DoclingParseDocumentBackend
from docling.document_converter import (
    FormatOption,
    DocumentConverter,
    PdfFormatOption,
    WordFormatOption,
    ExcelFormatOption,
    PowerpointFormatOption,
    MarkdownFormatOption,
    HTMLFormatOption,
    ImageFormatOption,
)
from docling.pipeline.simple_pipeline import SimplePipeline
from docling.pipeline.standard_pdf_pipeline import StandardPdfPipeline
from docling.datamodel.base_models import InputFormat, Table
from docling.datamodel.pipeline_options import PdfPipelineOptions

IMAGE_RESOLUTION_SCALE = 2.0

# Flask setup
app = Flask(__name__)

@app.route('/')
def home():
    """Render the home page."""
    return render_template('index.html')

def export_selected_formats(result, output_path, formats):
    """Export the document in selected formats."""
    output_path = Path(output_path)
    output_path.mkdir(parents=True, exist_ok=True)

    file_stem = result.input.file.stem

    if "md" in formats:
        with (output_path / f"{file_stem}.md").open("w") as fp:
            fp.write(result.document.export_to_markdown())
    if "txt" in formats:
        with (output_path / f"{file_stem}.txt").open("w", encoding="utf-8") as fp:
            fp.write(result.document.export_to_text())
    if "json" in formats:
        with (output_path / f"{file_stem}.json").open("w", encoding="utf-8") as fp:
            fp.write(json.dumps(result.document.export_to_dict(), indent=4))
    if "yaml" in formats:
        with (output_path / f"{file_stem}.yaml").open("w", encoding="utf-8") as fp:
            fp.write(yaml.safe_dump(result.document.export_to_dict(), default_flow_style=False, allow_unicode=True))

def convert_documents(input_paths, output_path, output_formats):
    """Convert documents and export in selected formats."""
    pipeline_options = PdfPipelineOptions()
    pipeline_options.images_scale = IMAGE_RESOLUTION_SCALE
    pipeline_options.generate_page_images = True
    pipeline_options.generate_picture_images = True

    doc_converter = DocumentConverter(
        allowed_formats=[InputFormat.XLSX, InputFormat.PDF, InputFormat.IMAGE, InputFormat.DOCX, InputFormat.HTML, InputFormat.PPTX, InputFormat.ASCIIDOC, InputFormat.MD],
        format_options={
            InputFormat.XLSX: FormatOption(pipeline_cls=SimplePipeline, backend=MsExcelDocumentBackend),
            InputFormat.DOCX: FormatOption(pipeline_cls=SimplePipeline, backend=MsWordDocumentBackend),
            InputFormat.PPTX: FormatOption(pipeline_cls=SimplePipeline, backend=MsPowerpointDocumentBackend),
            InputFormat.MD: FormatOption(pipeline_cls=SimplePipeline, backend=MarkdownDocumentBackend),
            InputFormat.ASCIIDOC: FormatOption(pipeline_cls=SimplePipeline, backend=AsciiDocBackend),
            InputFormat.HTML: FormatOption(pipeline_cls=SimplePipeline, backend=HTMLDocumentBackend),
            InputFormat.IMAGE: FormatOption(pipeline_cls=StandardPdfPipeline, backend=DoclingParseDocumentBackend),
            InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options),
        }
    )

    for input_path in input_paths:

        if not output_path.exists():
            output_path.mkdir(parents=True, exist_ok=True)

        conv_result = doc_converter.convert_all(input_paths)
        
        print(f"Converted all documents")
        for result in conv_result:
            # output_path = Path("output")
            print(
                f"Document {result.input.file.name} converted."
                f"\nSaved result output to: {str(output_path)}"
            )

            doc_filename = result.input.file.stem
            export_selected_formats(result=result, output_path=output_path, formats=output_formats)

            # Save images of figures and tables
            table_counter = 0
            picture_counter = 0

            # Create directory for saving images
            now = datetime.now()
            image_subdir = now.strftime(f"{doc_filename}_%Y-%m-%d_%H%M%S")
            image_dir = output_path / image_subdir

            if not image_dir.exists():
                image_dir.mkdir(parents=True, exist_ok=True)

            # Flags to check if we have any images or tables
            found_table_or_picture = False

            for element, _level in result.document.iterate_items():
                if isinstance(element, TableItem):
                    found_table_or_picture = True
                    table_counter += 1
                    element_image_filename = (
                        image_dir / f"{doc_filename}-table-{table_counter}.png"
                    )
                    
                    # Check if get_image() returns None before saving
                    image = element.get_image(result.document)
                    if image:
                        with element_image_filename.open("wb") as fp:
                            image.save(fp, "PNG")
                    else:
                        print(f"Warning: Table {table_counter} in {doc_filename} could not be converted to image.")

                if isinstance(element, PictureItem):
                    found_table_or_picture = True
                    picture_counter += 1
                    element_image_filename = (
                        image_dir / f"{doc_filename}-picture-{picture_counter}.png"
                    )
                    
                    # Check if get_image() returns None before saving
                    image = element.get_image(result.document)
                    if image:
                        with element_image_filename.open("wb") as fp:
                            image.save(fp, "PNG")
                    else:
                        print(f"Warning: Picture {picture_counter} in {doc_filename} could not be converted to image.")

            # If no table or picture was found, continue the program
            if not found_table_or_picture:
                print(f"No tables or images found in {doc_filename}. Skipping image save.")
                continue  # Skip to the next document

@app.route('/convert', methods=['POST'])
def convert():
    """API and web form endpoint for document conversion."""
    try:
        # Get form data (files, formats, and output path)
        files = request.files.getlist('files')
        output_formats = request.form.getlist('formats')
        output_path = request.form.get('output_path', 'output')

        if not files or not output_formats:
            return jsonify({"error": "Files and formats are required"}), 400

        # Save files to a temporary directory
        temp_dir = Path("temp_files")
        temp_dir.mkdir(parents=True, exist_ok=True)
        input_paths = []

        for file in files:
            file_path = temp_dir / file.filename
            file.save(file_path)
            input_paths.append(file_path)

        # Convert the documents
        convert_documents(input_paths, output_path, output_formats)

        # Return a response with download link(s)
        converted_files = []
        output_path = Path(output_path)
        for root, _, files in os.walk(output_path):
            for file in files:
                converted_files.append(f"{os.path.relpath(os.path.join(root, file), output_path)}")

        # For API requests
        if request.content_type == 'application/json':
            return jsonify({"message": "Documents converted successfully", "files": converted_files})

        # For web form requests
        return render_template('success.html', files=converted_files)

    except Exception as e:
        if request.content_type == 'application/json':
            return jsonify({"error": str(e)}), 500
        return f"<h1>Error: {str(e)}</h1>", 500


@app.route('/download/<path:filename>', methods=['GET'])
def download(filename):
    """API endpoint to download converted files."""
    return send_from_directory('output', filename)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=25041, debug=True)
