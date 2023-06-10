using EasyOffice.EasyExcel.Deprecate;
using EasyOffice.EasyExcel.Models;
using NPOI.XSSF.UserModel;

namespace EasyOffice.EasyExcel;

public sealed class ExportOption
{
    private XSSFWorkbook WorkBook { get; set; }

    public ExportOption(XSSFWorkbook workBook)
    {
        WorkBook = workBook;
    }
    
    /// <summary>
    /// Save Excel File in Specified Location and Returns File Path
    /// </summary>
    /// <param name="fileName">Name of the file</param>
    /// <param name="directoryPath">Directory where excel file need to save </param>
    /// <returns>File path of the saved excel file</returns>
    public string ExportIntoDirectory(string fileName,string directoryPath)
    {
        if (!Directory.Exists(directoryPath))
            Directory.CreateDirectory(directoryPath);
        
        var filePath = Path.Combine(directoryPath, GetFileName(fileName));
        
        var file = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        WorkBook.Write(file);
        file.Close();

        return filePath;
    }
    
    
    /// <summary>
    /// Write Excel File into MemoryStream
    /// </summary>
    /// <param name="fileName">Name of the file</param>
    /// <returns>Instance of FileExportModel</returns>
    public ExportData ExportIntoStream(string fileName)
    {
        var fileGenerateModel = new ExportData(fileName);
        
        WorkBook.Write(fileGenerateModel.Stream);

        return fileGenerateModel;
    }
    
    private static string GetFileName(string fileName)
    {
        if(fileName.Contains(ExcelConstant.SupportingExtension))
            return fileName;
        
        return fileName.Replace(".xls", "")+ ExcelConstant.SupportingExtension;
    }

    #region To be deprecated
    public FileGenerateModel ExportAsStream(string fileName)
    {
        var fileGenerateModel = new FileGenerateModel(fileName);
        
        WorkBook.Write(fileGenerateModel.Stream);

        return fileGenerateModel;
    }
    #endregion
}