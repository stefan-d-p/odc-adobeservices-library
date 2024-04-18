using OutSystems.ExternalLibraries.SDK;

namespace Without.Systems.AdobeServices.Structures;


[OSStructure(
    Description = "Single File Asset")]
public struct FileAsset
{
    [OSStructureField(
        Description = "Binary Data of File",
        DataType = OSDataType.BinaryData,
        IsMandatory = true)]
    public byte[] FileData;
    
    [OSStructureField(
        Description = "Filename of file including extension",
        DataType = OSDataType.Text,
        IsMandatory = true)]
    public string FileName;
    
    private FileAsset(byte[] fileData, string fileName)
    {
        FileData = fileData;
        FileName = fileName;
    }

    [OSIgnore]    
    public static FileAsset Create(byte[] fileData, string fileName)
    {
        return new FileAsset(fileData, fileName);
    }
}