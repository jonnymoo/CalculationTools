namespace ExcelService;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

public class Excel : IDisposable
{
    private SpreadsheetDocument? spreadsheetDocument = null;
    private WorksheetPart? wsPart = null;
    private MemoryStream? stream = null;
    
    public Excel(string spreadsheetString, string targetWorksheetName)
    {
        byte[] spreadsheetBytes = Convert.FromBase64String(spreadsheetString);
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

        wbPart.Workbook.CalculationProperties!.ForceFullCalculation = true;
        wbPart.Workbook.CalculationProperties!.FullCalculationOnLoad = true;

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

    public void SetInput(string cellName, string cellValue)
    {
           // Find the cell by its address using the cell name
        var cell = wsPart?.Worksheet?.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellName);

        if (cell != null)
        {
            // Update the cell value
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
            cell.CellValue = new CellValue(cellValue);

        }
        else
        {
            throw new Exception($"Cell '{cellName}' not found in input sheet'.");
        }
    }

    public byte[] Save()
    {
        // Save the modified spreadsheet 
        spreadsheetDocument!.Save();
        return stream!.ToArray();                 
    }
}
