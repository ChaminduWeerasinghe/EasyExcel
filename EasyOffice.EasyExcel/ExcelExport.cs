using System.Reflection;
using EasyOffice.EasyExcel.Attributes;
using EasyOffice.EasyExcel.Exceptions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace EasyOffice.EasyExcel;

public sealed class ExcelExport<TObj> where TObj : class
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
        var properties = type.GetProperties();
        if (properties.Length == 0)
            throw new PropertyInaccessibleException($"No properties found or properties are inaccessible in the provided class ({nameof(type)})");
        return properties.ToDictionary(propertyInfo => propertyInfo.Name, propertyInfo => 
            propertyInfo.GetCustomAttribute<HeaderName>()?.Name ?? propertyInfo.Name);
    }


    private static CellType GetCellTypeFromPropertyType(PropertyInfo propertyInfo)
    {
        if(propertyInfo.PropertyType == typeof(decimal) || propertyInfo.PropertyType == typeof(int) || propertyInfo.PropertyType == typeof(double))
            return CellType.Numeric;
        return propertyInfo.PropertyType == typeof(bool) ? CellType.Boolean : CellType.String;
    }

}

