using DocxToPdf.Api.Services;
using Microsoft.AspNetCore.Mvc;

namespace DocxToPdf.Api.Controllers;

[ApiController]
[Route("convert")]
public sealed class ConvertController : ControllerBase
{
    private readonly IWordToPdfConverter _converter;

    public ConvertController(IWordToPdfConverter converter)
    {
        _converter = converter;
    }

    [HttpPost]
    [Produces("application/pdf")]
    public async Task<IActionResult> Convert([FromForm] IFormFile file, CancellationToken cancellationToken)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest("Missing file");
        }

        var ext = Path.GetExtension(file.FileName);
        if (!string.Equals(ext, ".docx", StringComparison.OrdinalIgnoreCase))
        {
            return BadRequest("Only .docx supported");
        }

        string tempDocx = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);
        try
        {
            using (var fs = System.IO.File.Create(tempDocx))
            {
                await file.CopyToAsync(fs, cancellationToken);
            }

            byte[] pdfBytes = await _converter.ConvertDocxToPdfAsync(tempDocx, cancellationToken);
            return File(pdfBytes, "application/pdf", Path.ChangeExtension(file.FileName, ".pdf"));
        }
        finally
        {
            try { if (System.IO.File.Exists(tempDocx)) System.IO.File.Delete(tempDocx); } catch { }
        }
    }
}


