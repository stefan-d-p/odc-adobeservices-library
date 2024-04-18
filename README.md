# OutSystems Developer Cloud Adobe PDF Services External Logic Connector Library

This external logic library is based on the offical .net SDK by Adobe for their PDF Services. It wraps the following features

* **Create Document** - Convert Microsoft Office and Image file formats to PDF
* **Export Document** - Convert a PDF document to Microsoft Office or RTF file formats
* **OCR Document** - Perform optical character recognition (OCR) on a PDF and add a text layer to the document.
* **Linearize Document** - Optimize a PDF document for web view

## Actions

### CreateDocument

Create PDF document from Microsoft Office documents (Word, Excel and PowerPoint) and Image file formats.

**Input Parameters*

* `clientId` - Adobe Services Client Id
* `clientSecret` - Adobe Services Secret Key
* `fileAsset` - Source document binary data and filename

**Returns**

* `pdf` - Converted document binary data

### ExportDocument

Convert a PDF File to a non-PDF File

**Input Parameters**

* `clientId` - Adobe Services Client Id
* `clientSecret` - Adobe Services Secret Key
* `fileAsset` - Source document binary data and filename
* `targetFormat` - Format to convert the pdf document to. Allows values are doc, rtf, docx, pptx, xlsx

**Returns**

* `document` - Converted binary data document in target format

### OcrDocument

Perform Optical Character Recognition (OCR) on a PDF document

**Input Parameters**

* `clientId` - Adobe Services Client Id
* `clientSecret` - Adobe Services Secret Key
* `fileAsset` - Source document binary data and filename

**Returns**

* `pdf` - PDF document binary data with text layer

### LinearizeDocument

Linearize (optimize for Webview) a PDF document

**Input Parameters**

* `clientId` - Adobe Services Client Id
* `clientSecret` - Adobe Services Secret Key
* `fileAsset` - Source document binary data and filename

**Returns**

* `pdf` - Linearized PDF document binary data

### ImageDocument

Extracts all pages of a PDF document as images

**Input Parameters**

* `clientId` - Adobe Services Client Id
* `clientSecret` - Adobe Services Secret Key
* `fileAsset` - Source document binary data and filename
* `targetFormat` - Image format. Allowed values are png and jpeg

**Returns**

* `zip` - ZIP archive binary data containing the extracted images

# Run Tests

To run the nUnit tests you first have to provide your Adobe Services credentials as a secret

* Open a command prompt in the test project and run

```powershell
dotnet user-secrets init
dotnet user-secrets set ClientId "<your client id>"
dotnet user-secrets set ClientSecret "<your secret key>"
```

You can also set the DeepLAPIKey as an environment variable