using Microsoft.Extensions.Configuration;

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
        var result = _actions.CreateDocument(ClientId, ClientSecret, inputFile, "sampledoc.docx");
        Assert.That(result, Is.Not.Null);
    }

    [Test]
    public void Convert_PDF_ToMicrosoftWord()
    {
        byte[] inputFile = File.ReadAllBytes(@"docs\SamplePDF.pdf");
        var result = _actions.ExportDocument(ClientId, ClientSecret, inputFile, "samplepdf.pdf","docx");
        Assert.That(result, Is.Not.Null);
    }
}