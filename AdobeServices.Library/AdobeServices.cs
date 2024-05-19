using Adobe.PDFServicesSDK.auth;
using Adobe.PDFServicesSDK.core;
using Adobe.PDFServicesSDK.pdfops;
using Adobe.PDFServicesSDK.io;
using Adobe.PDFServicesSDK.options.compresspdf;
using Adobe.PDFServicesSDK.options.documentmerge;
using Adobe.PDFServicesSDK.options.exportpdf;
using Adobe.PDFServicesSDK.options.exportpdftoimages;
using Adobe.PDFServicesSDK.options.protectpdf;
using Newtonsoft.Json.Linq;
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

    /// <summary>
    /// Assemble Microsoft Word or PDF documents by merging JSON data with a Microsoft Word Template document
    /// </summary>
    /// <param name="clientId">Adobe Services Client Id</param>
    /// <param name="clientSecret">Adobe Services Secret Key</param>
    /// <param name="templateDocument">Microsoft Word Template document</param>
    /// <param name="jsonData">Serialized JSON data to merge with template</param>
    /// <param name="outputFormat">Specifies the output format. Supports docx or pdf</param>
    /// <returns>Binary Data of assembled document</returns>
    public byte[] GenerateDocument(string clientId, string clientSecret, FileAsset templateDocument, string jsonData, string outputFormat)
    {
        JObject jsonObject = JObject.Parse(jsonData);
        
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);

        DocumentMergeOptions options = new DocumentMergeOptions(jsonObject, ParseOutputFormat(outputFormat));
        DocumentMergeOperation documentMergeOperation = DocumentMergeOperation.CreateNew(options);
        
        using (MemoryStream templateDocumentStream = new MemoryStream(templateDocument.FileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef = FileRef.CreateFromStream(templateDocumentStream,
                MimeTypesMap.GetMimeType(templateDocument.FileName));
            documentMergeOperation.SetInput(inputDocumentRef);
            FileRef outputDocumentRef = documentMergeOperation.Execute(executionContext);
            outputDocumentRef.SaveAs(targetDocumentStream);
            return targetDocumentStream.ToArray();
        }
    }

    public byte[] ExportDocument(string clientId, string clientSecret, FileAsset fileAsset, string targetFormat)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);
        
        ExportPDFOperation exportPdfOperation = ExportPDFOperation.CreateNew(ParseExportTargetFormat(targetFormat));
        
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
        ExportPDFToImagesOperation  exportPdfToImagesOperation = ExportPDFToImagesOperation.CreateNew(ParseExportPdfImagesTargetFormat(targetFormat));
        
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

    public byte[] ProtectDocument(string clientId, string clientSecret, FileAsset fileAsset, ProtectDocumentOptions options)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);
        
        var builder = ProtectPDFOptions.PasswordProtectOptionsBuilder();
        if (!string.IsNullOrEmpty(options.EncryptionAlgorithm))
        {
            builder.SetEncryptionAlgorithm(ParseEncryptionAlgorithm(options.EncryptionAlgorithm));
        }
        else
        {
            throw new ArgumentException("Encryption Algorithm must be specified");
        }

        if (!string.IsNullOrEmpty(options.OwnerPassword))
        {
            builder.SetOwnerPassword(options.OwnerPassword);
        }
        else
        {
            throw new ArgumentException("Owner password must be specified");
        }

        if (!string.IsNullOrEmpty(options.UserPassword))
            builder.SetUserPassword(options.UserPassword);
        if (!string.IsNullOrEmpty(options.ContentEncryption))
        {
            builder.SetContentEncryption(ParseContentEncryption(options.ContentEncryption));
        }
       
        if (options.Permissions != null)
            builder.SetPermissions(ParsePermissions(options.Permissions.ToArray()));
        ProtectPDFOptions protectDocumentOptions = builder.Build();

        ProtectPDFOperation protectPdfOperation = ProtectPDFOperation.CreateNew(protectDocumentOptions);
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileAsset.FileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef = FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileAsset.FileName));
            protectPdfOperation.SetInput(inputDocumentRef);
            FileRef outputDocumentRefs = protectPdfOperation.Execute(executionContext);
            outputDocumentRefs.SaveAs(targetDocumentStream);
            return targetDocumentStream.ToArray();
        }
    }


    public byte[] UnprotectDocument(string clientId, string clientSecret, FileAsset fileAsset, string password)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);
        
        RemoveProtectionOperation removeProtectionOperation = RemoveProtectionOperation.CreateNew();
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileAsset.FileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef = FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileAsset.FileName));
            removeProtectionOperation.SetInput(inputDocumentRef);
            removeProtectionOperation.SetPassword(password);
            FileRef outputDocumentRef = removeProtectionOperation.Execute(executionContext);
            outputDocumentRef.SaveAs(targetDocumentStream);
            return targetDocumentStream.ToArray();
        }
    }

    public byte[] CompressDocument(string clientId, string clientSecret, FileAsset fileAsset)
    {
        ExecutionContext executionContext = CreateExecutionContext(clientId, clientSecret);
        CompressPDFOperation compressPdfOperation = CompressPDFOperation.CreateNew();
        
        using (MemoryStream sourceDocumentStream = new MemoryStream(fileAsset.FileData))
        using (MemoryStream targetDocumentStream = new MemoryStream())
        {
            FileRef inputDocumentRef = FileRef.CreateFromStream(sourceDocumentStream, MimeTypesMap.GetMimeType(fileAsset.FileName));
            compressPdfOperation.SetInput(inputDocumentRef);
            FileRef outputDocumentRef = compressPdfOperation.Execute(executionContext);
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

    private OutputFormat ParseOutputFormat(string outputFormat)
    {
        switch (outputFormat.ToLower())
        {
            case "docx":
                return OutputFormat.DOCX;
            case "pdf":
                return OutputFormat.PDF;
            default:
                throw new ArgumentException($"Invalid output format of {outputFormat}");
        }
    }

    private ExportPDFTargetFormat ParseExportTargetFormat(string targetFormat)
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
    
    private ExportPDFToImagesTargetFormat ParseExportPdfImagesTargetFormat(string targetFormat)
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

    private ContentEncryption ParseContentEncryption(string contentEncryption)
    {
        switch (contentEncryption.ToUpper())
        {
            case "ALL_CONTENT":
                return ContentEncryption.ALL_CONTENT;
            case "ALL_CONTENT_EXCEPT_METADATA":
                return ContentEncryption.ALL_CONTENT_EXCEPT_METADATA;
            default:
                throw new ArgumentException($"Invalid content encryption of {contentEncryption}");
        }
    }

    private EncryptionAlgorithm ParseEncryptionAlgorithm(string encryptionAlgorithm)
    {
        switch (encryptionAlgorithm.ToUpper())
        {
            case "AES_128":
                return EncryptionAlgorithm.AES_128;
            case "AES_256":
                return EncryptionAlgorithm.AES_256;
            default:
                throw new ArgumentException($"Invalid encryption algorithm of {encryptionAlgorithm}");
        }
    }
    
    private Permissions? ParsePermissions(string[] permissions)
    {
        if (permissions.Length == 0) return null;

        Permissions pdfPermissions = Permissions.CreateNew();

        foreach (string permission in permissions)
        {
            pdfPermissions.AddPermission(ParsePermission(permission));
        }

        return pdfPermissions;
    }

    private Permission ParsePermission(string permission)
    {
        switch (permission.ToUpper()) {
            case "COPY_CONTENT":
                return Permission.COPY_CONTENT;
            case "EDIT_CONTENT":
                return Permission.EDIT_CONTENT;
            case "EDIT_ANNOTATIONS":
                return Permission.EDIT_ANNOTATIONS;
            case "PRINT_LOW_QUALITY":
                return Permission.PRINT_LOW_QUALITY;
            case "EDIT_DOCUMENT_ASSEMBLY":
                return Permission.EDIT_DOCUMENT_ASSEMBLY;
            case "PRINT_HIGH_QUALITY":
                return Permission.PRINT_HIGH_QUALITY;
            case "EDIT_FILL_AND_SIGN_FORM_FIELDS":
                return Permission.EDIT_FILL_AND_SIGN_FORM_FIELDS;
            default:
                throw new ArgumentException($"Invalid permission of {permission}");
        }
    }
    
    
}