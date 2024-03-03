using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Threading.Tasks;
using Newtonsoft.Json;
using ExcelService;

namespace CalculationTools;
public class RunExcelCalc
{
    private readonly ILogger _logger;

    public RunExcelCalc(ILoggerFactory loggerFactory)
    {
        _logger = loggerFactory.CreateLogger<RunExcelCalc>();
    }

    [Function("RunExcelCalc")]
    public HttpResponseData Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
    {
        _logger.LogInformation("C# HTTP trigger function called.");

        // Parse JSON input
        var jsonString = new StreamReader(req.Body).ReadToEnd();
        var inputData = JsonConvert.DeserializeObject<dynamic>(jsonString);

        // Validate input data (replace with your validation logic)
        if (inputData == null || !inputData?.ContainsKey("spreadsheet") || !inputData?.ContainsKey("inputs") || !inputData?.ContainsKey("inputSheetName"))
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
            string spreadsheetString = inputData!.spreadsheet;

            // Retrieve the target worksheet name
            string targetWorksheetName = inputData.inputSheetName;

            using Excel excel = new(spreadsheetString, targetWorksheetName);

            // Input mapping
            foreach (var item in inputData.inputs)
            {
                string cellName = item.CellName;
                string cellValue = item.Value;
                excel.SetCellValue(cellName, cellValue);
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
        _logger.LogInformation("C# HTTP trigger function called.");

        // Parse JSON input
        var jsonString = new StreamReader(req.Body).ReadToEnd();
        var inputData = JsonConvert.DeserializeObject<dynamic>(jsonString);

        // Validate input data (replace with your validation logic)
        if (inputData == null || !inputData?.ContainsKey("spreadsheet") || !inputData?.ContainsKey("values") || !inputData?.ContainsKey("outputSheetName"))
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
            string spreadsheetString = inputData!.spreadsheet;

            // Retrieve the target worksheet name
            string targetWorksheetName = inputData.outputSheetName;

            using Excel excel = new(spreadsheetString, targetWorksheetName);

            var jsonArray = new List<Dictionary<string, string>>();

            // Ouput mapping
            foreach (var item in inputData.values)
            {
                string cellName = item.CellName;
                string output = excel.GetCellValue(cellName);
                _logger.LogInformation($"Cell {cellName}, Value {output}");

                // Create a dictionary object for each cell
                var cellData = new Dictionary<string, string>();
                cellData.Add("CellName", cellName);
                cellData.Add("Value", output);

                // Add the dictionary to the JSON array
                jsonArray.Add(cellData);
            }

            // Convert the JSON array to a string
            string outputJsonString = JsonConvert.SerializeObject(jsonArray);


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
