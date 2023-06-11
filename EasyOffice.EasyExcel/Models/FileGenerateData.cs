namespace EasyOffice.EasyExcel.Models;

public class FileGenerateData
{
    public FileGenerateData(string fileName)
    {
        FileName = fileName;
        Stream = new MemoryStream();
    }
    public MemoryStream Stream { get;}
    public string FileName { get;}
}