namespace ExcelService;

using OfficeOpenXml;
using System.IO;

public class Excel : IDisposable
{
    ExcelPackage? package = null;
    ExcelWorksheet? worksheet = null;
    private MemoryStream? stream = null;

    static Excel()
    {
        // Set the license context to non-commercial
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
    
    public Excel(string spreadsheetString) : this(Convert.FromBase64String(spreadsheetString))
    {
    }

    public Excel(byte[] spreadsheetBytes)
    {
        stream = new MemoryStream(spreadsheetBytes);

        package = new ExcelPackage(stream);

    }

    public void SetCurrentSheet(string targetWorksheetName)
    {
        if(worksheet != null)
        {
            worksheet.Dispose();
            worksheet = null;
        }
        
        worksheet = package?.Workbook.Worksheets.FirstOrDefault(s => s.Name == targetWorksheetName);

        if (worksheet == null)
        {
            // Handle the case where the worksheet is not found
            throw new Exception($"Sheet '{targetWorksheetName}' not found in the spreadsheet.");
        }
    }


    public void Dispose()
    {
        if(worksheet != null) 
        {
            worksheet.Dispose();
        }

        if(package != null)
        {
            package.Dispose();
        }

        if(stream != null)
        {
            stream.Dispose();
        }
    }

    public void SetCellValue(string cellName, string cellValue)
    {
        var cell = GetCell(cellName);
        // Update the cell value
        cell.SetCellValue(0,0,cellValue);
    }

    public string GetCellValue(string cellName)
    {
        var cell = GetCell(cellName);
        return cell.GetCellValue<string>();
    }

    public byte[] Save()
    {
        // Save the modified spreadsheet 
        package?.Workbook.Calculate();
        using MemoryStream calculatedStream = new();
        package?.SaveAs(calculatedStream);
        return calculatedStream.ToArray();                 
    }

    private ExcelRangeBase GetCell(string cellName)
    {
        // Find the cell by its address using the cell name
        var cell = worksheet?.Cells.FirstOrDefault(c => c.Address == cellName);

        if (cell != null)
        {
            return cell;
        }
        else
        {
            throw new Exception($"Cell '{cellName}' not found in sheet");
        }
    }
}
