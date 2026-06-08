from django.shortcuts import render
from django.http import JsonResponse, FileResponse
from django.views.decorators.csrf import csrf_exempt
import tempfile
import os
import re
import zipfile
from io import BytesIO

from .utils import (
    read_excel,
    generate_qr,
    extract_invoice_number,
    insert_qr_into_docx
)


def home(request):
    return render(request, "index.html")


@csrf_exempt
def upload_files(request):

    if request.method == "POST":

        data_file = request.FILES.get('data_file')
        templates = request.FILES.getlist('templates')

        # Read Excel mapping
        excel_map = read_excel(data_file)

        # Remove old PDFs from output folder
        output_folder = "media/output"

        if os.path.exists(output_folder):

            for file in os.listdir(output_folder):

                if file.lower().endswith(".pdf"):

                    os.remove(
                        os.path.join(output_folder, file)
                    )

        results = []

        for template in templates:

            try:

                print(f"📄 Processing: {template.name}")

                # Save uploaded DOCX temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    tmp.write(template.read())
                    temp_docx_path = tmp.name

                # STEP 1 → Try filename extraction
                match = re.search(r'Invoice_(\d+)', template.name)

                if match:
                    invoice_no = match.group(1)
                    print("✅ Invoice from filename:", invoice_no)

                else:
                    invoice_no = extract_invoice_number(temp_docx_path)

                # STEP 2 → Validate invoice
                if invoice_no is None:

                    results.append({
                        "file": template.name,
                        "status": "Invoice not found"
                    })

                    continue

                # STEP 3 → Match Excel data
                data = excel_map.get(invoice_no)

                if data is None:

                    results.append({
                        "file": template.name,
                        "status": f"No data for invoice {invoice_no}"
                    })

                    continue

                irn = data["irn"]
                ack_no = data["ack_no"]
                ack_date = data["ack_date"]

                # STEP 4 → Generate QR
                qr_image = generate_qr(
                    irn,
                    ack_no,
                    ack_date
                )

                # STEP 5 → Insert QR into DOCX
                final_docx = insert_qr_into_docx(
                    temp_docx_path,
                    qr_image,
                    irn,
                    ack_no,
                    ack_date,
                    template.name
                )

                results.append({
                    "file": template.name,
                    "status": "Success",
                    "docx": os.path.basename(final_docx)
                })

            except Exception as e:

                results.append({
                    "file": template.name,
                    "status": f"Error: {str(e)}"
                })

        return JsonResponse({
            "results": results
        })


# -----------------------------------
# DOWNLOAD FILE
# -----------------------------------

def download_file(request, filename):

    file_path = os.path.join("media/output", filename)

    if not os.path.exists(file_path):

        return JsonResponse({
            "error": "File not found"
        })

    return FileResponse(
        open(file_path, 'rb'),
        as_attachment=True
    )

def download_all(request):

    output_folder = "media/output"

    zip_buffer = BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zip_file:

        for filename in os.listdir(output_folder):

            # Only include DOCX files
            if not filename.lower().endswith(".docx"):
                continue

            file_path = os.path.join(
                output_folder,
                filename
            )

            if os.path.isfile(file_path):

                zip_file.write(
                    file_path,
                    arcname=filename
                )

    zip_buffer.seek(0)

    return FileResponse(
        zip_buffer,
        as_attachment=True,
        filename="Processed_Invoices.zip"
    )