using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Threading.Tasks;
using Newtonsoft.Json;
using ExcelService;
using System.Web;

namespace CalculationTools;
public class RunExcelCalc
{
    private readonly ILogger _logger;

    public RunExcelCalc(ILoggerFactory loggerFactory)
    {
        _logger = loggerFactory.CreateLogger<RunExcelCalc>();
    }

    [Function("SetValues")]
    public HttpResponseData Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
    {
        // Parse JSON input
        var jsonString = new StreamReader(req.Body).ReadToEnd();

        var inputData = JsonConvert.DeserializeObject<dynamic>(jsonString);

        // Validate input data (replace with your validation logic)
        if (inputData == null || !inputData?.ContainsKey("SpreadSheet") || !inputData?.ContainsKey("Inputs") )
        {
            var response = req.CreateResponse(HttpStatusCode.BadGateway);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");
            response.WriteString("Invalid input data format");
            return response;
        }

        // Process the data
        try
        {
            // Decode base64-encoded spreadsheet
            string spreadsheetString = inputData!.SpreadSheet;

            spreadsheetString = spreadsheetString.Replace("&#13;&#10;","");
            
            using Excel excel = new(spreadsheetString);

            // Input mapping
            foreach (var sheet in inputData.Inputs)
            {
                // Retrieve the target worksheet name
                string targetWorksheetName = sheet.Name;
                var cells = sheet.Value;
                excel.SetCurrentSheet(targetWorksheetName);

                foreach(var cell in cells)
                {
                    string cellName = cell.Name;
                    string cellValue = cell.Value;
                    excel.SetCellValue(cellName, cellValue);
                }
            }

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/octet-stream");
            response.WriteBytes(excel.Save());
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "An error occurred processing the Excel file.");
            var response = req.CreateResponse(HttpStatusCode.BadGateway);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");
            response.WriteString("Error processing data");
            return response;
        }
    } 


    [Function("GetValues")]
    public HttpResponseData GetValues([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
    {
        // Parse JSON input
        var jsonString = new StreamReader(req.Body).ReadToEnd();
        var inputData = JsonConvert.DeserializeObject<dynamic>(jsonString);

        // Validate input data (replace with your validation logic)
        if (inputData == null || !inputData?.ContainsKey("SpreadSheet") || !inputData?.ContainsKey("Outputs"))
        {
            var response = req.CreateResponse(HttpStatusCode.BadGateway);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");
            response.WriteString("Invalid input data format");
            return response;
        }

        // Process the data
        try
        {
            // Decode base64-encoded spreadsheet
            string spreadsheetString = inputData!.SpreadSheet;

            spreadsheetString = spreadsheetString.Replace("&#13;&#10;","");
            
            using Excel excel = new(spreadsheetString);

            // Ouput mapping
            var sheetData = new Dictionary<string, Dictionary<string, string>>();
            foreach (var sheet in inputData.Outputs)
            {
                // Retrieve the target worksheet name
                string targetWorksheetName = sheet.Name;
                var cells = sheet.Value;
                excel.SetCurrentSheet(targetWorksheetName);

                var cellData = new Dictionary<string, string>();
                sheetData.Add(targetWorksheetName, cellData);

                foreach(var cell in cells)
                {
                    string cellName = cell.Name;
                    string output = excel.GetCellValue(cellName);
                    cellData.Add(cell.Name, output);
                }
            }

           string outputJsonString = JsonConvert.SerializeObject(sheetData);


            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/json");
            response.WriteString(outputJsonString);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "An error occurred processing the Excel file.");
            var response = req.CreateResponse(HttpStatusCode.BadGateway);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");
            response.WriteString("Error processing data");
            return response;
        }
    } 
}
