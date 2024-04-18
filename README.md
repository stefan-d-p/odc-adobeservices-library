# OutSystems Developer Cloud Adobe PDF Services External Logic Connector Library

This external logic library is based on the offical .net SDK by Adobe for their PDF Services. It wraps the following features

* **Create Document** - Convert Microsoft Office and Image file formats to PDF
* **Export Document** - Convert a PDF document to Microsoft Office or RTF file formats
* **OCR Document** - Perform optical character recognition (OCR) on a PDF and add a text layer to the document.
* **Linearize Document** - Optimize a PDF document for web view

# Run Tests

To run the nUnit tests you first have to provide your Adobe Services credentials as a secret

* Open a command prompt in the test project and run

```powershell
dotnet user-secrets init
dotnet user-secrets set ClientId "<your client id>"
dotnet user-secrets set ClientSecret "<your secret key>"
```

You can also set the DeepLAPIKey as an environment variable