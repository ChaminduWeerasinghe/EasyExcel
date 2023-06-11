using EasyOffice.EasyExcel.Exceptions;
using Microsoft.AspNetCore.Http;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.SS.UserModel;
using FileNotFoundException = EasyOffice.EasyExcel.Exceptions.FileNotFoundException;

namespace EasyOffice.EasyExcel.Models;

public class ImportOption
{
    /// <summary>
    /// Import From IFormFile
    /// </summary>
    /// <param name="file"></param>
    /// <exception cref="Exceptions.FileNotFoundException"></exception>
    public static ImportOption ImportFrom(IFormFile file)
    {
        if (file is not {Length: > 0}) 
            throw new EmptyFileException("No data contained in the file");
        
        using var memoryStream = new MemoryStream();
        file.CopyToAsync(memoryStream).Wait();
        return new ImportOption(memoryStream);

    }
    
    /// <summary>
    /// Import From Stream
    /// </summary>
    /// <param name="file"></param>
    /// <exception cref="Exceptions.FileNotFoundException"></exception>
    public static ImportOption ImportFrom(Stream file)
    {
        if (file is not {Length: > 0}) 
            throw new EmptyFileException("No data contained in the stream");
        
        using var memoryStream = new MemoryStream();
        file.CopyToAsync(memoryStream).Wait();
        return new ImportOption(memoryStream);

    }
    
    /// <summary>
    /// Import From FilePath
    /// </summary>
    /// <param name="filePath"></param>
    /// <exception cref="Exceptions.FileNotFoundException"></exception>
    
    public static ImportOption ImportFrom(string filePath)
    {
        if (!File.Exists(filePath)) 
            throw new FileNotFoundException("Invalid file path");
        
        if (new FileInfo(filePath).Length <= 0) 
            throw new EmptyFileException("No data contained in the file");
        
        using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        return new ImportOption(fileStream);
    }

    private ImportOption(Stream inputStream)
    {
        try
        {
            Workbook = WorkbookFactory.Create(inputStream);
        }
        catch (InvalidFormatException)
        {
            throw new InvalidFileTypeException($"Only {ExcelConstant.SupportingExtension} files are supported");
        }
    }
    
    public IWorkbook Workbook { get; }
}