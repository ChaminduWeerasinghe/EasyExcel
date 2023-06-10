namespace EasyOffice.EasyExcel.Models;

public class ExportData
{
    public ExportData(string fileName)
    {
        FileName = fileName;
        Stream = new MemoryStream();
    }
    public MemoryStream Stream { get;}
    public string FileName { get;}
}