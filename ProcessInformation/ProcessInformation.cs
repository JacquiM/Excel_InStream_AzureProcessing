using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Collections.Generic;
using Azure.Storage.Blobs;
using ClosedXML.Excel;

namespace ProcessInformation
{
    public static class ProcessInformation
    {
        [FunctionName("ProcessInformation")]
        public static async Task Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request to process information.");

            try
            {
                // Create blank request body
                string requestBody = String.Empty;

                // Populate request body
                using (StreamReader streamReader = new StreamReader(req.Body))
                {
                    requestBody = await streamReader.ReadToEndAsync();
                }

                log.LogInformation("Request body received.");

                // Convert JSON string to object
                var data = JsonConvert.DeserializeObject<InputData>(requestBody.ToString());

                log.LogInformation("JSON deserialised and parsed through model.");
                
                // Create new list of PersonalData and populate with InputData
                List<PersonalDetails> personalDetailsList = data.PersonalDetails;

                log.LogInformation("New list of personal details created.");

                // Save Excel Template file from Blob Storage
                BlobStorage blobStorage = new BlobStorage();
                var stream = blobStorage.ReadTemplateFromBlob();

                log.LogInformation("Template retrieved from blob.");

                // Populate Excel Spreadsheet
                ExcelProcessing excelProcessing = new ExcelProcessing();

                var excelStream = excelProcessing.PopulateExcelTable(personalDetailsList, stream, "Details", "PersonalDetails");

                log.LogInformation("Excel populated");

                // Create response
                var response = req.HttpContext.Response;
                response.ContentType = "application/json";

                // Create stream writer and memory stream
                using StreamWriter streamWriter = new StreamWriter(response.Body);

                // Add the memory stream to the stream writer/request.Body
                await streamWriter.WriteLineAsync("[\"" + Convert.ToBase64String(excelStream.ToArray()) + "\"]");
                await streamWriter.FlushAsync();
            }
            catch (Exception e)
            {
                log.LogInformation($"Error processing information: {e.Message}");
            }
        }
    }
}

/* Helpers */
class BlobStorage
{
    public MemoryStream ReadTemplateFromBlob()
    {
        // TO DO: Add Connection String and move into secure storage
        string connectionString = "<connection string>";
        string container = "excel-templates";

        // Get blob
        BlobContainerClient blobContainerClient = new BlobContainerClient(connectionString, container);
        var blob = blobContainerClient.GetBlobClient("Information.xlsx");

        MemoryStream memoryStream = new MemoryStream();

        // Download blob
        blob.DownloadTo(memoryStream);

        return memoryStream;
    }
}

class ExcelProcessing
{
    public MemoryStream PopulateExcelTable (List<PersonalDetails> personalDetailsList, MemoryStream stream, string sheetName, string tableName)
    {
        // Link to existing workbook
        using var wbook = new XLWorkbook(stream);

        // Link to sheer
        var ws = wbook.Worksheet(sheetName);

        // Set row offset
        int currentRow = 2;
        
        // Write each record to sheet - Uncomment below to write into an Excel document without tables
        //foreach(var item in personalDetailsList)
        //{
        //    ws.Cell(currentRow, 1).Value = item.Name;
        //    ws.Cell(currentRow, 2).Value = item.Surname;
        //    ws.Cell(currentRow, 3).Value = item.DateOfBirth;
        //}

        // Replace data in existing Excel Table
        var table = wbook.Table(tableName);
        table.ReplaceData(personalDetailsList, propagateExtraColumns: true);

        // Save file
        wbook.SaveAs(stream);

        // Create new stream
        MemoryStream memoryStream = stream;

        // Return stream
        return memoryStream;
    }
}

/* Models */
partial class PersonalDetails
{
    [JsonProperty("Name")]
    public string Name { get; set; }
    [JsonProperty("Surname")]
    public string Surname { get; set; }
    [JsonProperty("DateOfBirth")]
    public string DateOfBirth { get; set; }
}

partial class InputData
{
    [JsonProperty("PersonalDetails")]
    public List<PersonalDetails> PersonalDetails { get; set; }
}