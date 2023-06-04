namespace EasyOffice.EasyExcel.Deprecate;

[Obsolete("FileGenerateModel is deprecated, please use FileExportModel instead.")]
public class FileGenerateModel
{
    public FileGenerateModel(string fileName)
    {
        FileName = fileName;
        Stream = new MemoryStream();
    }
    public MemoryStream Stream { get;}
    public string FileName { get;}
}