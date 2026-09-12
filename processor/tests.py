from io import BytesIO

from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from openpyxl import Workbook

from .utils import build_qr_payload, generate_qr, read_excel


class SecurityAndUploadTests(TestCase):
    def test_download_file_rejects_path_traversal(self):
        response = self.client.get('/download/../../secret.txt/')

        self.assertEqual(response.status_code, 400)
        self.assertIn('invalid', response.json()['error'].lower())

    def test_upload_rejects_unsupported_file_types(self):
        invalid_excel = SimpleUploadedFile(
            'bad.txt',
            b'not an excel file',
            content_type='text/plain'
        )

        valid_template = SimpleUploadedFile(
            'invoice.docx',
            b'not a real docx',
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

        response = self.client.post(
            '/upload/',
            {
                'data_file': invalid_excel,
                'templates': [valid_template],
            },
            format='multipart'
        )

        self.assertEqual(response.status_code, 400)
        self.assertIn('excel', response.json()['error'].lower())

    def test_read_excel_mapping_and_qr_include_doc_metadata(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Sheet1'
        sheet.append(['doc no', 'irn', 'ack no', 'ack date', 'doc typ', 'doc date', 'inv value.'])
        sheet.append(['103202501371', 'IRN123', 'ACK123', '2026-03-10 22:42:00', 'Invoice', '2026-03-10', '1500.00'])

        stream = BytesIO()
        workbook.save(stream)
        stream.seek(0)

        mapping = read_excel(stream)

        self.assertIn('103202501371', mapping)
        self.assertEqual(mapping['103202501371']['doc_type'], 'Invoice')
        self.assertEqual(mapping['103202501371']['doc_date'], '2026-03-10')
        self.assertEqual(mapping['103202501371']['inv_value'], '1500.00')

        payload = build_qr_payload(
            mapping['103202501371']['irn'],
            mapping['103202501371']['ack_no'],
            mapping['103202501371']['ack_date'],
            mapping['103202501371']['doc_no'],
            mapping['103202501371']['doc_date'],
            mapping['103202501371']['inv_value'],
        )

        self.assertIn('Doc No. : 103202501371', payload)
        self.assertIn('Doc Date : 2026-03-10', payload)
        self.assertIn('Inv Value : 1500.00', payload)
