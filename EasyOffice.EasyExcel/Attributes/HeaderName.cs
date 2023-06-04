namespace EasyOffice.EasyExcel.Attributes;

public sealed class HeaderName: Attribute
{
    public string Name { get; }

    public HeaderName(string name)
    {
        Name = name;
    }
    
}