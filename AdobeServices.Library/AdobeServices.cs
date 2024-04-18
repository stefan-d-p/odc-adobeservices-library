using Adobe.PDFServicesSDK.auth;
using Adobe.PDFServicesSDK.core;
using Adobe.PDFServicesSDK.pdfops;
using Adobe.PDFServicesSDK.io;
using Adobe.PDFServicesSDK.options.exportpdf;
using ExecutionContext = Adobe.PDFServicesSDK.ExecutionContext;

namespace Without.Systems.AdobeServices;

public class AdobeServices : IAdobeServices
{
    /// <summary>
    /// Create PDF document from Microsoft Office documents (Word, Excel and PowerPoint) and Image file formats.
    /// </summary>
    /// <param name="clientId">Adobe Services Client Id</param>
    /// <param name="clientSecret">Adobe Services Secret Key</param>
    /// <param name="fileData">Source document binary data</param>
    /// <param name="fileName">Source document filename including extension</param>
    /// <returns>Binary data of PDF document</returns>
    public byte[] CreateDocument(string clientId,
        string clientSecret, byte[] fileData, string fileName)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);
        
        CreatePDFOperation createPdfOperation = CreatePDFOperation.CreateNew();
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef = FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileName));
            createPdfOperation.SetInput(inputDocumentRef);
            
            FileRef outputDocumentRef = createPdfOperation.Execute(executionContext);
            
            outputDocumentRef.SaveAs(targetDocumentStream);
            
            return targetDocumentStream.ToArray();
        }
    }

    public byte[] ExportDocument(string clientId, string clientSecret, byte[] fileData, string fileName, string targetFormat)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);
        
        ExportPDFOperation exportPdfOperation = ExportPDFOperation.CreateNew(GetExportTargetFormat(targetFormat));
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef = FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileName));
            exportPdfOperation.SetInput(inputDocumentRef);
 
            FileRef outputDocumentRef = exportPdfOperation.Execute(executionContext);
            outputDocumentRef.SaveAs(targetDocumentStream);
            return targetDocumentStream.ToArray();
        }
    }

    public byte[] OcrDocument(string clientId, string clientSecret, byte[] fileData, string fileName)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);

        OCROperation ocrOperation = OCROperation.CreateNew();
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef = FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileName));
            ocrOperation.SetInput(inputDocumentRef);
            FileRef outputDocumentRef = ocrOperation.Execute(executionContext);
            outputDocumentRef.SaveAs(targetDocumentStream);
            return targetDocumentStream.ToArray();
        }
    }

    public byte[] LinearizeDocument(string clientId, string clientSecret, byte[] fileData, string fileName)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);
        
        LinearizePDFOperation linearizePdfOperation = LinearizePDFOperation.CreateNew();
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef =
                FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileName));
            linearizePdfOperation.SetInput(inputDocumentRef);
            FileRef outputDocumentRef = linearizePdfOperation.Execute(executionContext);
            outputDocumentRef.SaveAs(targetDocumentStream);
            return targetDocumentStream.ToArray();
        }
    }

    /// <summary>
    /// Build Execution context with credentials
    /// </summary>
    /// <param name="clientId">Adobe Service Client Id</param>
    /// <param name="clientSecret">Adobe Service Client Secret</param>
    /// <returns></returns>
    private ExecutionContext CreateExecutionContext(string clientId, string clientSecret)
    {
        Credentials credentials = Credentials.ServicePrincipalCredentialsBuilder()
            .WithClientId(clientId)
            .WithClientSecret(clientSecret)
            .Build();
        
        return ExecutionContext.Create(credentials);
    }

    private ExportPDFTargetFormat GetExportTargetFormat(string targetFormat)
    {
        switch (targetFormat.ToLower()) {
            case "doc":
                return ExportPDFTargetFormat.DOC;
            case "rtf":
                return ExportPDFTargetFormat.RTF;
            case "docx":
                return ExportPDFTargetFormat.DOCX;
            case "pptx":
                return ExportPDFTargetFormat.PPTX;
            case "xlsx":
                return ExportPDFTargetFormat.XLSX;
            default:
                throw new ArgumentException($"Invalid target format of {targetFormat}");
        }
    }
}