using System.Reflection;
using EasyOffice.EasyExcel.Attributes;
using EasyOffice.EasyExcel.Exceptions;
using NPOI.SS.UserModel;
using ImportOption = EasyOffice.EasyExcel.Models.ImportOption;

namespace EasyOffice.EasyExcel;

public sealed class ExcelImport<TObj> where TObj : class
{
    /// <summary>
    /// Read xlsx file and map data to List of TObject. Throws exception when try to map nullable value to not nullable property
    /// </summary>
    /// <param name="importOption"></param>
    /// <param name="sheetName"></param>
    /// <exception cref="InvalidValueException">When try to map blank value from excel into not nullable property in TObj</exception>
    /// <returns>List of TObj</returns>
    public static List<TObj> ReadExcel(ImportOption importOption,string sheetName) 
        => RetrieveExcelData(importOption.Workbook, sheetName,true);
    
    
    /// <summary>
    /// Read xlsx file and map data to List of TObject. Set default value when try to map nullable to not nullable property
    /// </summary>
    /// <param name="importOption"></param>
    /// <param name="sheetName"></param>
    /// <returns>List of TObj</returns>
    public static List<TObj> ReadExcelInSafe(ImportOption importOption,string sheetName) 
        => RetrieveExcelData(importOption.Workbook, sheetName);
    

    private static List<TObj> RetrieveExcelData(IWorkbook workbook, string sheetName, bool isThrowOnInvalidData = false)
    {
        
        var sheet = workbook.GetSheet(sheetName);
        var rows = sheet.GetRowEnumerator();
        var mapping = new Dictionary<int, PropertyInfo>();
        var insertObjectList = new List<TObj>();

        while (rows.MoveNext())
        {
            IRow curRow;
            try
            {
                curRow = (IRow)rows.Current!;
            }
            catch (Exception)
            {
                continue;
            }
            
            if (mapping.Count > 0)
            {
                var insertObject = Activator.CreateInstance<TObj>();
                foreach (var cell in curRow.Cells.Where(cell => cell.CellType is not CellType.Blank))
                {
                    if (!mapping.TryGetValue(cell.ColumnIndex, out var propertyInfo)) continue;

                    var cellValue = cell.ToString();

                    var value = ConvertValue(propertyInfo, cellValue, cell.Address.FormatAsString(),isThrowOnInvalidData);
                    
                    SetValueForProperty(propertyInfo,insertObject,value);
                }
                insertObjectList.Add(insertObject);
            }
            else
            {
                mapping =  GetRowIndexPropertyInfoMapping(sheet.GetRow(curRow.RowNum));
            }
        }

        return insertObjectList;
    }

    private static void SetValueForProperty(PropertyInfo propertyInfo, TObj insertObject, object? value)
    {
        try
        {
            propertyInfo.SetValue(insertObject,value);
        }
        catch (ArgumentException e)
        {
            throw new PropertyInaccessibleException($"Error while setting value to property {propertyInfo.Name} in {typeof(TObj).Name}",e);
        }
    }
    
    private static Dictionary<int, PropertyInfo> GetRowIndexPropertyInfoMapping(IRow row)
    {
        var map = new Dictionary<int, PropertyInfo>();
        var properties = typeof(TObj).GetProperties().ToList();
        foreach (var cell in row.Cells)
        {
            var cellValue = cell.ToString();

            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                foreach (var property in properties)
                {
                    var propertyName = property.Name;
                    var propertyAttributeName = property.GetCustomAttribute<HeaderName>()?.Name;
                    
                    if (string.Equals(cellValue, propertyName, StringComparison.CurrentCultureIgnoreCase) ||
                        string.Equals(cellValue, propertyAttributeName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        map.Add(cell.ColumnIndex, property);
                    }
                }
            }
        }

        return map;
    }
    

    private static object? ConvertValue(PropertyInfo propertyInfo, string? cellValue, string cellName, bool isThrowOnInvalidData)
    {
        var propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType);
        
        if(isThrowOnInvalidData && propertyType is null && string.IsNullOrWhiteSpace(cellValue))
            throw new InvalidValueException($"Value of property '{propertyInfo.Name}' in '{typeof(TObj)}' cannot be null (Cell Address : {cellName})");
        
        var propertyInfoPropertyType = propertyType ?? propertyInfo.PropertyType;
        return string.IsNullOrWhiteSpace(cellValue) ? null : Convert.ChangeType(cellValue, propertyInfoPropertyType);
    }
}