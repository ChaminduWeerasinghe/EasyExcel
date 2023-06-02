using System.Reflection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace EasyOffice.EasyExcel;

public class ExcelGenerate<TObj> where TObj : class
{
    
    /// <summary>
    /// Do basic excel creation features
    /// </summary>
    /// <param name="sheetName">Name of the sheet that expected in excel</param>
    /// <param name="models">List of data that insert into excel</param>
    /// <param name="excludes">Columns which exclude from </param>
    /// <returns>ExportOption</returns>
    public static ExportOption GenerateExcel(string sheetName,List<TObj> models,List<string>? excludes = null)
    {
        
        var workBook = new XSSFWorkbook();
        var sheet = workBook.CreateSheet(sheetName);

        var propertyType = models.First().GetType();
        var propertyNames = GetPropertyNamesAndAnnotationsByClass(propertyType);// PropertyName, Annotation
        
       
        if (excludes is {Count: > 0})
            propertyNames = propertyNames.Where(x => !excludes.Contains(x.Key) || !excludes.Contains(x.Value))
                .ToDictionary(x => x.Key, x => x.Value);
        
        var columnIdNameMap = CreateHeadersInSheet(sheet,propertyNames);
        
        DataInsertion(sheet, models, propertyType, columnIdNameMap);
        
        return new ExportOption(workBook);
    }
    
    
    private static void DataInsertion(ISheet sheet,List<TObj> models,Type type,Dictionary<int,string> columnIdNameMap)
    {
        var rowIndex = 1;
        foreach (var model in models)
        {
            CreateRowInSheet(sheet, model, type, columnIdNameMap, rowIndex);
            rowIndex++;
        }
        foreach (var key in columnIdNameMap.Keys)
        {
            sheet.AutoSizeColumn(key);
        }
    }
    
    private static void CreateRowInSheet(ISheet sheet, TObj model,Type type,Dictionary<int,string> columnIdNameMap , int rowIndex)
    {
        var row = sheet.CreateRow(rowIndex);
        foreach (var propertyMap in columnIdNameMap)
        {
            var propertyType = type.GetProperty(propertyMap.Value);
            var value = propertyType!.GetValue(model)?.ToString();
            value = string.IsNullOrWhiteSpace(value) ? string.Empty : value;
            var cell = row.CreateCell(propertyMap.Key);
            cell.SetCellType(GetCellTypeFromPropertyType(propertyType));
            cell.SetCellValue(value);
        }
    }

    /// <summary>
    /// Create Header in Excel Sheet with Column Names from Annotation
    /// </summary>
    /// <param name="sheet">Sheet that extracted from workbook</param>
    /// <param name="propertyNames">Dictionary which have property name and annotation</param>
    /// <returns>Dictionary which have Column Index and PropertyName</returns>
    private static Dictionary<int, string> CreateHeadersInSheet(ISheet sheet, Dictionary<string,string> propertyNames)
    {
        var row = sheet.CreateRow(0);
        var cellIndex = 0;
        var columnIdNameMap = new Dictionary<int, string>();

        for (var index = 0; index < propertyNames.Count; index++)
        {
            var column = propertyNames.ElementAt(index).Key;
            var columnNameAttribute = propertyNames.ElementAt(index).Value;
            
            columnIdNameMap.Add(cellIndex, column);
            var cell = row.CreateCell(cellIndex);
            cell.SetCellValue(columnNameAttribute);
            cell.RichStringCellValue.ApplyFont(0, columnNameAttribute.Length, new XSSFFont
            {
                IsBold = true
            });
            cellIndex++;
        }

        return columnIdNameMap;
    }
    
    /// <summary>
    /// Returns Dictionary of PropertyName and Annotation
    /// </summary>
    /// <param name="type">Type of Class that Property Details are Extracted From</param>
    /// <returns>Dictionary which have property name and annotation</returns>
    private static Dictionary<string,string> GetPropertyNamesAndAnnotationsByClass(Type type)
    {
        return type.GetProperties().ToDictionary(propertyInfo => propertyInfo.Name, propertyInfo => 
            propertyInfo.GetCustomAttribute<HeaderName>()?.Name ?? propertyInfo.Name);
    }


    private static CellType GetCellTypeFromPropertyType(PropertyInfo propertyInfo)
    {
        if(propertyInfo.PropertyType == typeof(decimal) || propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(double))
            return CellType.Numeric;
        return propertyInfo.PropertyType == typeof(bool) ? CellType.Boolean : CellType.String;
    }
}

public class ExportOption
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
    public string ExportIntoFileLocation(string fileName,string directoryPath)
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
    /// <returns>Instance of FileGenerateModel</returns>
    public FileGenerateModel ExportAsStream(string fileName)
    {
        var fileGenerateModel = new FileGenerateModel(fileName);
        
        WorkBook.Write(fileGenerateModel.Stream);

        return fileGenerateModel;
    }
    
    private static string GetFileName(string fileName)
    {
        if(fileName.Contains(ExcelExportConst.Extension))
            return fileName;
        
        return fileName.Replace(".xls", "")+ ExcelExportConst.Extension;
        
    }
    
}


public class HeaderName: Attribute
{
    public string Name { get; }

    public HeaderName(string name)
    {
        Name = name;
    }
    
}


public class ExcelExportConst
{
    public const string Extension = ".xlsx";
    public const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    public const string ExcelFileName = "ExcelFileName";
    public const string ExcelFilePath = "Workbooks";
}