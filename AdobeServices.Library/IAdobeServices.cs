using System.ComponentModel.DataAnnotations;
using OutSystems.ExternalLibraries.SDK;
using Without.Systems.AdobeServices.Structures;

namespace Without.Systems.AdobeServices
{
    [OSInterface(
        Name = "AdobePDFServices",
        Description = "Adobe Acrobat Services is a cloud-based platform that enables developers to integrate PDF functionality into their applications. With this component, developers can create, convert, and edit PDFs, as well as extract data from PDF files",
        IconResourceName = "Without.Systems.AdobeServices.Resources.AdobeServices.png")]
    public interface IAdobeServices
    {
        [OSAction(
            Description =
                "Create PDF document from Microsoft Office documents (Word, Excel and PowerPoint) and Image file formats.",
            ReturnName = "pdf",
            ReturnType = OSDataType.BinaryData,
            ReturnDescription = "Converted document",
            IconResourceName = "Without.Systems.AdobeServices.Resources.CreateDocument.png")]
        byte[] CreateDocument(
            [OSParameter(
                Description = "Adobe Services Client Id",
                DataType = OSDataType.Text)]
            string clientId,
            [OSParameter(
                Description = "Adobe Services Client Secret",
                DataType = OSDataType.Text)]
            string clientSecret,
            [OSParameter(
                Description = "Source document and filename",
                DataType = OSDataType.InferredFromDotNetType)]
            FileAsset fileAsset);

        [OSAction(
            Description =
                "Assemble a Microsoft Word or PDF document by merging data with a Microsoft Word Template Document",
            ReturnName = "result",
            ReturnType = OSDataType.BinaryData,
            ReturnDescription = "Assembled Document",
            IconResourceName = "Without.Systems.AdobeServices.Resources.GenerateDocument.png")]
        byte[] GenerateDocument(
            [OSParameter(
                Description = "Adobe Services Client Id",
                DataType = OSDataType.Text)]
            string clientId,
            [OSParameter(
                Description = "Adobe Services Client Secret",
                DataType = OSDataType.Text)]
            string clientSecret,
            [OSParameter(
                Description = "Microsoft Word Template Document",
                DataType = OSDataType.BinaryData)]
            FileAsset templateDocument, 
            [OSParameter(
                Description = "Serialized JSON data to merge with template document",
                DataType = OSDataType.Text)]
            string jsonData,
            [OSParameter(
                Description = "Output format of the generated document. Supports docx and pdf als values.",
                DataType = OSDataType.Text)]
            string outputFormat);

        [OSAction(
            Description = "Convert a PDF File to a non-PDF File",
            ReturnName = "document",
            ReturnType = OSDataType.BinaryData,
            ReturnDescription = "Converted document",
            IconResourceName = "Without.Systems.AdobeServices.Resources.ExportDocument.png")]
        byte[] ExportDocument([OSParameter(
                Description = "Adobe Services Client Id",
                DataType = OSDataType.Text)]
            string clientId,
            [OSParameter(
                Description = "Adobe Services Client Secret",
                DataType = OSDataType.Text)]
            string clientSecret,
            [OSParameter(
                Description = "Source document and filename",
                DataType = OSDataType.InferredFromDotNetType)]
            FileAsset fileAsset,
            [OSParameter(
                Description = "Format to convert the pdf document to. Allows values are doc, rtf, docx, pptx, xlsx",
                DataType = OSDataType.Text)]
            string targetFormat);
        
        [OSAction(
            Description = "Perform Optical Character Recognition (OCR) on a PDF document",
            ReturnName = "pdf",
            ReturnType = OSDataType.BinaryData,
            ReturnDescription = "PDF document with text layer",
            IconResourceName = "Without.Systems.AdobeServices.Resources.OcrDocument.png")]
        byte[] OcrDocument([OSParameter(
                Description = "Adobe Services Client Id",
                DataType = OSDataType.Text)]
            string clientId,
            [OSParameter(
                Description = "Adobe Services Client Secret",
                DataType = OSDataType.Text)]
            string clientSecret,
            [OSParameter(
                Description = "Source document and filename",
                DataType = OSDataType.InferredFromDotNetType)]
            FileAsset fileAsset);
        
        [OSAction(
            Description = "Linearize (optimize for Webview) a PDF document",
            ReturnName = "pdf",
            ReturnType = OSDataType.BinaryData,
            ReturnDescription = "Linearized PDF document",
            IconResourceName = "Without.Systems.AdobeServices.Resources.LinearizeDocument.png")]
        byte[] LinearizeDocument([OSParameter(
                Description = "Adobe Services Client Id",
                DataType = OSDataType.Text)]
            string clientId,
            [OSParameter(
                Description = "Adobe Services Client Secret",
                DataType = OSDataType.Text)]
            string clientSecret,
            [OSParameter(
                Description = "Source document and filename",
                DataType = OSDataType.InferredFromDotNetType)]
            FileAsset fileAsset);

        [OSAction(
            Description = "Convert all pages of a PDF document to images",
            ReturnName = "zip",
            ReturnType = OSDataType.BinaryData,
            ReturnDescription = "ZIP archive containing the image files",
            IconResourceName = "Without.Systems.AdobeServices.Resources.ImageDocument.png")]
        byte[] ImageDocument(
            [OSParameter(
                Description = "Adobe Services Client Id",
                DataType = OSDataType.Text)]
            string clientId,
            [OSParameter(
                Description = "Adobe Services Client Secret",
                DataType = OSDataType.Text)]
            string clientSecret,
            [OSParameter(
                Description = "Source document and filename",
                DataType = OSDataType.InferredFromDotNetType)]
            FileAsset fileAsset,
            [OSParameter(
                Description = "Target format of images. Allows values are png and jpeg",
                DataType = OSDataType.Text)]
            string targetFormat);

        [OSAction(
            Description = "Protect a PDF Document",
            ReturnName = "pdf",
            ReturnType = OSDataType.BinaryData,
            ReturnDescription = "Protected PDF Document",
            IconResourceName = "Without.Systems.AdobeServices.Resources.ProtectDocument.png")]
        byte[] ProtectDocument(
            [OSParameter(
                Description = "Adobe Services Client Id",
                DataType = OSDataType.Text)]
            string clientId,
            [OSParameter(
                Description = "Adobe Services Client Secret",
                DataType = OSDataType.Text)]
            string clientSecret,
            [OSParameter(
                Description = "Source document and filename",
                DataType = OSDataType.InferredFromDotNetType)]
            FileAsset fileAsset,
            [OSParameter(
                Description = "Protect Document Options")]
            ProtectDocumentOptions options);

        [OSAction(
            Description = "Remove protection from a PDF document",
            ReturnName = "pdf",
            ReturnType = OSDataType.BinaryData,
            ReturnDescription = "Unprotected PDF Document",
            IconResourceName = "Without.Systems.AdobeServices.Resources.UnprotectDocument.png")]
        byte[] UnprotectDocument(
            [OSParameter(
                Description = "Adobe Services Client Id",
                DataType = OSDataType.Text)]
            string clientId,
            [OSParameter(
                Description = "Adobe Services Client Secret",
                DataType = OSDataType.Text)]
            string clientSecret,
            [OSParameter(
                Description = "Source document and filename",
                DataType = OSDataType.InferredFromDotNetType)]
            FileAsset fileAsset,
            [OSParameter(
                Description = "Password to unprotect the document",
                DataType = OSDataType.Text)]
            string password);

        [OSAction(
            Description = "Compress a PDF Document",
            ReturnName = "pdf",
            ReturnType = OSDataType.BinaryData,
            ReturnDescription = "Compressed PDF Document",
            IconResourceName = "Without.Systems.AdobeServices.Resources.CompressDocument.png")]
        public byte[] CompressDocument(
            [OSParameter(
                Description = "Adobe Services Client Id",
                DataType = OSDataType.Text)]
            string clientId,
            [OSParameter(
                Description = "Adobe Services Client Secret",
                DataType = OSDataType.Text)]
            string clientSecret,
            [OSParameter(
                Description = "Source document and filename",
                DataType = OSDataType.InferredFromDotNetType)]
            FileAsset fileAsset);
    }
}