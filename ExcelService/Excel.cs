namespace ExcelService;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using System.IO;

public class Excel : IDisposable
{
    private SpreadsheetDocument? spreadsheetDocument = null;
    private WorksheetPart? wsPart = null;
    private MemoryStream? stream = null;

    static Excel()
    {
        // Set the license context to non-commercial
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
    
    public Excel(string spreadsheetString, string targetWorksheetName) : this(Convert.FromBase64String(spreadsheetString), targetWorksheetName)
    {
    }

    public Excel(byte[] spreadsheetBytes, string targetWorksheetName)
    {
        stream = new MemoryStream(spreadsheetBytes);

        spreadsheetDocument = SpreadsheetDocument.Open(stream, true);

        // Retrieve a reference to the workbook part.
        WorkbookPart? wbPart = spreadsheetDocument.WorkbookPart;

        // Access the worksheet using the name
        var worksheet = wbPart!.Workbook.Descendants<Sheet>()
                        .FirstOrDefault(s => s.Name == targetWorksheetName);

        if (worksheet == null)
        {
            // Handle the case where the worksheet is not found
            throw new Exception($"Sheet '{targetWorksheetName}' not found in the spreadsheet.");
        }

        // Retrieve a reference to the worksheet part.
        wsPart = (WorksheetPart)wbPart!.GetPartById(worksheet.Id!);
    }


    public void Dispose()
    {
        if(spreadsheetDocument != null) 
        {
            spreadsheetDocument.Dispose();
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
        cell.DataType = new EnumValue<CellValues>(CellValues.String);
        cell.CellValue = new CellValue(cellValue);
    }

    public string GetCellValue(string cellName)
    {
        var cell = GetCell(cellName);
        if(cell.CellValue == null)
        {
            throw new Exception($"Cell '{cellName}' value not set");
        }
        return cell.CellValue.Text;
    }

    public byte[] Save()
    {
        // Save the modified spreadsheet 
        spreadsheetDocument!.Save();
        using ExcelPackage package = new ExcelPackage(stream);
        package.Workbook.Calculate();
        using MemoryStream calculatedStream = new();
        package.SaveAs(calculatedStream);
        return calculatedStream.ToArray();                 
    }

    private Cell GetCell(string cellName)
    {
           // Find the cell by its address using the cell name
        var cell = wsPart?.Worksheet?.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellName);

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
