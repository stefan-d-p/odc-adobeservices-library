using OutSystems.ExternalLibraries.SDK;

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
                Description = "Source document to convert to PDF",
                DataType = OSDataType.BinaryData)]
            byte[] fileData,
            [OSParameter(
                Description = "Filename of source document including extension",
                DataType = OSDataType.Text)]
            string fileName);

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
                Description = "Source PDF document to convert",
                DataType = OSDataType.BinaryData)]
            byte[] fileData,
            [OSParameter(
                Description = "Filename of source PDF document including pdf extension",
                DataType = OSDataType.Text)]
            string fileName,
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
                Description = "Source PDF document to perform OCR on",
                DataType = OSDataType.BinaryData)]
            byte[] fileData,
            [OSParameter(
                Description = "Filename of source PDF document including pdf extension",
                DataType = OSDataType.Text)]
            string fileName);
        
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
                Description = "Source PDF document to linearize",
                DataType = OSDataType.BinaryData)]
            byte[] fileData,
            [OSParameter(
                Description = "Filename of source PDF document including pdf extension",
                DataType = OSDataType.Text)]
            string fileName);

        [OSAction(
            Description = "Convert all pages of a PDF document to images",
            ReturnName = "zip",
            ReturnType = OSDataType.BinaryData,
            ReturnDescription = "ZIP archive containing the image files",
            IconResourceName = "Without.Systems.AdobeServices.Resources.ImageDocument.png")]
        byte[] ImageDocument([OSParameter(
                Description = "Adobe Services Client Id",
                DataType = OSDataType.Text)]
            string clientId,
            [OSParameter(
                Description = "Adobe Services Client Secret",
                DataType = OSDataType.Text)]
            string clientSecret,
            [OSParameter(
                Description = "Source PDF document for page imaging",
                DataType = OSDataType.BinaryData)]
            byte[] fileData,
            [OSParameter(
                Description = "Filename of source PDF document including pdf extension",
                DataType = OSDataType.Text)]
            string fileName,
            [OSParameter(
                Description = "Target format of images. Allows values are png and jpeg",
                DataType = OSDataType.Text)]
            string targetFormat);
    }
}