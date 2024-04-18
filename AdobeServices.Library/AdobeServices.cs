using Adobe.PDFServicesSDK.auth;
using Adobe.PDFServicesSDK.core;
using Adobe.PDFServicesSDK.pdfops;
using Adobe.PDFServicesSDK.io;
using Adobe.PDFServicesSDK.options.exportpdf;
using Adobe.PDFServicesSDK.options.exportpdftoimages;
using Without.Systems.AdobeServices.Structures;
using ExecutionContext = Adobe.PDFServicesSDK.ExecutionContext;

namespace Without.Systems.AdobeServices;

public class AdobeServices : IAdobeServices
{
    /// <summary>
    /// Create PDF document from Microsoft Office documents (Word, Excel and PowerPoint) and Image file formats.
    /// </summary>
    /// <param name="clientId">Adobe Services Client Id</param>
    /// <param name="clientSecret">Adobe Services Secret Key</param>
    /// <param name="fileAsset">Source document binary data and filename</param>
    /// <returns>Binary data of PDF document</returns>
    public byte[] CreateDocument(string clientId,
        string clientSecret, FileAsset fileAsset)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);
        
        CreatePDFOperation createPdfOperation = CreatePDFOperation.CreateNew();
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileAsset.FileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef = FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileAsset.FileName));
            createPdfOperation.SetInput(inputDocumentRef);
            
            FileRef outputDocumentRef = createPdfOperation.Execute(executionContext);
            
            outputDocumentRef.SaveAs(targetDocumentStream);
            
            return targetDocumentStream.ToArray();
        }
    }

    public byte[] ExportDocument(string clientId, string clientSecret, FileAsset fileAsset, string targetFormat)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);
        
        ExportPDFOperation exportPdfOperation = ExportPDFOperation.CreateNew(GetExportTargetFormat(targetFormat));
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileAsset.FileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef = FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileAsset.FileName));
            exportPdfOperation.SetInput(inputDocumentRef);
 
            FileRef outputDocumentRef = exportPdfOperation.Execute(executionContext);
            outputDocumentRef.SaveAs(targetDocumentStream);
            return targetDocumentStream.ToArray();
        }
    }

    public byte[] OcrDocument(string clientId, string clientSecret, FileAsset fileAsset)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);

        OCROperation ocrOperation = OCROperation.CreateNew();
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileAsset.FileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef = FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileAsset.FileName));
            ocrOperation.SetInput(inputDocumentRef);
            FileRef outputDocumentRef = ocrOperation.Execute(executionContext);
            outputDocumentRef.SaveAs(targetDocumentStream);
            return targetDocumentStream.ToArray();
        }
    }

    public byte[] LinearizeDocument(string clientId, string clientSecret, FileAsset fileAsset)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);
        
        LinearizePDFOperation linearizePdfOperation = LinearizePDFOperation.CreateNew();
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileAsset.FileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef =
                FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileAsset.FileName));
            linearizePdfOperation.SetInput(inputDocumentRef);
            FileRef outputDocumentRef = linearizePdfOperation.Execute(executionContext);
            outputDocumentRef.SaveAs(targetDocumentStream);
            return targetDocumentStream.ToArray();
        }
    }

    public byte[] ImageDocument(string clientId, string clientSecret, FileAsset fileAsset,
        string targetFormat)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);
        ExportPDFToImagesOperation  exportPdfToImagesOperation = ExportPDFToImagesOperation.CreateNew(GetExportPDFImagesTargetFormat(targetFormat));
        
        exportPdfToImagesOperation.SetOutputType(ExportPDFToImagesOutputType.ZIP_OF_IMAGES);
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileAsset.FileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef = FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileAsset.FileName));
            exportPdfToImagesOperation.SetInput(inputDocumentRef);
            List<FileRef> outputDocumentRefs = exportPdfToImagesOperation.Execute(executionContext);
            outputDocumentRefs.First().SaveAs(targetDocumentStream);
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
    
    private ExportPDFToImagesTargetFormat GetExportPDFImagesTargetFormat(string targetFormat)
    {
        switch (targetFormat.ToLower())
        {
            case "png":
                return ExportPDFToImagesTargetFormat.PNG;
            case "jpeg":
                return ExportPDFToImagesTargetFormat.JPEG;
            default:
                throw new ArgumentException($"Invalid target format of {targetFormat}");
        }
    }
}