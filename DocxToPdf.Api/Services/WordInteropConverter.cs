using Microsoft.Office.Interop.Word;

namespace DocxToPdf.Api.Services;

public sealed class WordInteropConverter : IWordToPdfConverter
{
    public async Task<byte[]> ConvertDocxToPdfAsync(string inputDocxPath, CancellationToken cancellationToken)
    {
        // Interop is blocking; run on thread pool to avoid blocking request thread
        return await Task.Run(() => ConvertInternal(inputDocxPath), cancellationToken).ConfigureAwait(false);
    }

    private static byte[] ConvertInternal(string inputDocxPath)
    {
        Application? wordApp = null;
        Document? document = null;
        string? pdfTempPath = null;

        try
        {
            wordApp = new Application
            {
                Visible = false,
                ScreenUpdating = false,
                DisplayAlerts = WdAlertLevel.wdAlertsNone
            };

            object missing = System.Reflection.Missing.Value;
            object readOnly = true;
            object isVisible = false;
            object fileName = inputDocxPath;

            document = wordApp.Documents.Open(ref fileName, ReadOnly: ref readOnly, Visible: ref isVisible);

            pdfTempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");

            document.ExportAsFixedFormat(
                pdfTempPath,
                WdExportFormat.wdExportFormatPDF,
                OpenAfterExport: false,
                OptimizeFor: WdExportOptimizeFor.wdExportOptimizeForPrint,
                Range: WdExportRange.wdExportAllDocument,
                From: 0,
                To: 0,
                Item: WdExportItem.wdExportDocumentContent,
                IncludeDocProps: true,
                KeepIRM: true,
                CreateBookmarks: WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                DocStructureTags: true,
                BitmapMissingFonts: true,
                UseISO19005_1: false);

            byte[] bytes = File.ReadAllBytes(pdfTempPath);
            return bytes;
        }
        finally
        {
            try
            {
                if (document != null)
                {
                    document.Close(WdSaveOptions.wdDoNotSaveChanges);
                }
            }
            catch { }

            try
            {
                if (wordApp != null)
                {
                    wordApp.Quit(WdSaveOptions.wdDoNotSaveChanges);
                }
            }
            catch { }

            if (pdfTempPath != null && File.Exists(pdfTempPath))
            {
                try { File.Delete(pdfTempPath); } catch { }
            }
        }
    }
}


