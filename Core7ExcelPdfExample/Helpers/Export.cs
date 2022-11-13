using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace Core7ExcelPdfExample.Helpers;

public class Export
{
    public static byte[] Excel<T>(List<T> table, string filename)
    {
        using ExcelPackage pack = new ExcelPackage();
        ExcelWorksheet ws = pack.Workbook.Worksheets.Add(filename);
        ws.Cells["A1"].LoadFromCollection(table, true, TableStyles.Light1);
        return pack.GetAsByteArray();
    }
}
