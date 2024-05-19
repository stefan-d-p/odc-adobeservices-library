using Microsoft.Extensions.Configuration;
using Without.Systems.AdobeServices.Structures;

namespace Without.Systems.AdobeServices.Test;

public class Tests
{
    private static readonly IAdobeServices _actions = new AdobeServices();
    private string ClientId;
    private string ClientSecret;

    [SetUp]
    public void Setup()
    {
        IConfiguration configuration = new ConfigurationBuilder()
            .AddUserSecrets<Tests>()
            .AddEnvironmentVariables()
            .Build();
        
        ClientId = configuration["ClientId"] ?? throw new InvalidOperationException();
        ClientSecret = configuration["ClientSecret"] ?? throw new InvalidOperationException();
    }

    [Test]
    public void Convert_MicrosoftWord_ToPDF()
    {
        byte[] inputFile = File.ReadAllBytes(@"docs\SampleDOC.docx");
        var result = _actions.CreateDocument(ClientId, ClientSecret, FileAsset.Create(inputFile, "sampledoc.docx"));
        Assert.That(result, Is.Not.Null);
    }

    [Test]
    public void Convert_PDF_ToMicrosoftWord()
    {
        byte[] inputFile = File.ReadAllBytes(@"docs\SamplePDF.pdf");
        var result = _actions.ExportDocument(ClientId, ClientSecret, FileAsset.Create(inputFile, "samplepdf.pdf"), "docx");
        Assert.That(result, Is.Not.Null);
    }

    [Test]
    public void Convert_PDF_ToImages()
    {
        byte[] inputFile = File.ReadAllBytes(@"docs\SamplePDF.pdf");
        var result = _actions.ImageDocument(ClientId, ClientSecret, FileAsset.Create(inputFile, "samplepdf.pdf"), "png");
        Assert.That(result, Is.Not.Null);
    }

    [Test]
    public void Protect_PDF_OwnerPassword()
    {
        byte[] inputFile = File.ReadAllBytes(@"docs\SamplePDF.pdf");
        ProtectDocumentOptions options = new ProtectDocumentOptions
        {
            OwnerPassword = "testpassword",
            ContentEncryption = "ALL_CONTENT",
            EncryptionAlgorithm = "AES_128"
        };
        
        var result = _actions.ProtectDocument(ClientId, ClientSecret, FileAsset.Create(inputFile, "samplepdf.pdf"), options);
        Assert.That(result, Is.Not.Null);
    }
    
    [Test]
    public void Protect_PDF_UserPassword()
    {
        byte[] inputFile = File.ReadAllBytes(@"docs\SamplePDF.pdf");
        ProtectDocumentOptions options = new ProtectDocumentOptions
        {
            OwnerPassword = "testpassword",
            UserPassword = "openpassword",
            ContentEncryption = "ALL_CONTENT",
            EncryptionAlgorithm = "AES_128"
        };
        
        var result = _actions.ProtectDocument(ClientId, ClientSecret, FileAsset.Create(inputFile, "samplepdf.pdf"), options);
        Assert.That(result, Is.Not.Null);
    }

    [Test]
    public void Compress_PDF_Document()
    {
        byte[] inputFile = File.ReadAllBytes(@"docs\SamplePDF.pdf");
        var result = _actions.CompressDocument(ClientId, ClientSecret, FileAsset.Create(inputFile, "samplepdf.pdf"));
        Assert.That(result, Is.Not.Null);
    }
}