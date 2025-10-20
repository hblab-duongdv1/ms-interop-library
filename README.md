## DOCX to PDF API (Microsoft Office Interop) in Windows Container

Important notes:
- This API uses Microsoft.Office.Interop.Word and requires Microsoft Word to be installed in the Windows container or host. Running Office inside containers is generally not supported by Microsoft and may violate licensing/EULA. Proceed only if you understand and accept the risks and have proper licensing.
- For production and Linux portability, prefer server-side libraries like LibreOffice or commercial SDKs that do not require Office.

### Build and Run (Windows containers)

1) Switch Docker to Windows containers.
2) Build the image:

```bash
docker build -t interop-docx2pdf:win .
```

3) Run the container (ensure an image with Word installed is used, or modify the Dockerfile to include your base with Office):

```bash
docker run --rm -p 8080:8080 --name docx2pdf interop-docx2pdf:win
```

### Convert with curl

```bash
curl -X POST "http://localhost:8080/convert" \
  -H "Accept: application/pdf" \
  -F "file=@/absolute/path/to/sample.docx" \
  --output output.pdf
```

### Endpoint
- POST `/convert`: multipart/form-data with field `file` (.docx). Returns PDF bytes.

### Troubleshooting
- If you see COM or Word activation failures, Word is likely not installed/activated in the container.
- Interop requires the Windows Desktop COM automation model; run with elevated permissions and single-instance Word.


