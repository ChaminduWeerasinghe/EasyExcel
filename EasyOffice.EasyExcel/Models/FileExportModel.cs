namespace EasyOffice.EasyExcel.Models;

public class FileExportModel
{
    public FileExportModel(string fileName)
    {
        FileName = fileName;
        Stream = new MemoryStream();
    }
    public MemoryStream Stream { get; set; }
    public string FileName { get; set; }
}