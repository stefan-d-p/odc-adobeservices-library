using OutSystems.ExternalLibraries.SDK;

namespace Without.Systems.AdobeServices.Structures;

[OSStructure(
    Description = "PDF Document Protection options")]
public struct ProtectDocumentOptions
{
    [OSStructureField(
        Description = "Document Permissions. Supports the values COPY_CONTENT, EDIT_CONTENT, EDIT_ANNOTATIONS, PRINT_LOW_QUALITY, PRINT_HIGH_QUALITY,EDIT_DOCUMENT_ASSEMBLY and EDIT_FILL_AND_SIGN_FORM_FIELDS",
        DataType = OSDataType.Text,
        IsMandatory = false)]
    public List<string>? Permissions;
    
    [OSStructureField(
        Description = "Owner Password",
        DataType = OSDataType.Text,
        IsMandatory = true)]
    public string OwnerPassword;
    
    [OSStructureField(
        Description = "User Password",
        DataType = OSDataType.Text,
        IsMandatory = false)]
    public string? UserPassword;

    [OSStructureField(
        Description = "Content Encryption. Default is ALL_CONTENT.",
        DataType = OSDataType.Text,
        IsMandatory = false,
        DefaultValue = "ALL_CONTENT")]
    public string ContentEncryption;
    
    [OSStructureField(
        Description = "Encryption Algorithm. Supported values are AES_128 and AES_256",
        DataType = OSDataType.Text,
        IsMandatory = true)]
    public string EncryptionAlgorithm;
}