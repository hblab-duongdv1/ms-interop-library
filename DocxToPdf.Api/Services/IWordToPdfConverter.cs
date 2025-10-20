namespace DocxToPdf.Api.Services;

public interface IWordToPdfConverter
{
    Task<byte[]> ConvertDocxToPdfAsync(string inputDocxPath, CancellationToken cancellationToken);
}


