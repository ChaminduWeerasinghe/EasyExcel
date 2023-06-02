namespace EasyOffice.EasyExcel;

public class FileGenerateModel
{
    public FileGenerateModel(string fileName)
    {
        FileName = fileName;
        Stream = new MemoryStream();
    }
    public MemoryStream Stream { get; set; }
    public string FileName { get; set; }
}