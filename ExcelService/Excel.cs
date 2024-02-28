namespace ExcelService;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


public class Excel : IDisposable
{
    private SpreadsheetDocument? spreadsheetDocument = null;
    private Sheet? worksheet = null;
    private MemoryStream? stream = null;
    
    public Excel(string spreadsheetString, string targetWorksheetName)
    {
        byte[] spreadsheetBytes = Convert.FromBase64String(spreadsheetString);
        stream = new MemoryStream(spreadsheetBytes);

        spreadsheetDocument = SpreadsheetDocument.Open(stream, true);

        // Access the worksheet using the name
        worksheet = spreadsheetDocument?.WorkbookPart?.Workbook.Descendants<Sheet>()
                        .FirstOrDefault(s => s.Name == targetWorksheetName);

        if (worksheet == null)
        {
            // Handle the case where the worksheet is not found
            throw new Exception($"Sheet '{targetWorksheetName}' not found in the spreadsheet.");
        }

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
        var cell = worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellName);

        if (cell != null)
        {
            // Update the cell value
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
        using var memoryStream = new MemoryStream();
        spreadsheetDocument!.WorkbookPart!.Workbook.Save(memoryStream);
        return memoryStream.ToArray();                 
    }
}
