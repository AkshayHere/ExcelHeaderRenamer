using Microsoft.Extensions.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Serilog;
using System.Text.Json;

class Program
{
    static void Main(string[] args)
    {
        IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        // Setup Serilog
        Log.Logger = new LoggerConfiguration()
            .ReadFrom.Configuration(config)
            .CreateLogger();

        // Project root
        //string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory)!
        //    .Parent!
        //    .Parent!
        //    .Parent!
        //    .FullName;
        string projectRoot = AppContext.BaseDirectory;

        // Assets folder
        string assetsFolder = Path.Combine(projectRoot, "Assets");

        string xmlInputFilePath = config["FileSettings:XmlInputFilePath"] ?? throw new Exception("XmlInputFilePath not configured");
        string xmlOutputFilePath = config["FileSettings:XmlOutputFilePath"] ?? throw new Exception("XmlOutputFilePath not configured");
        string xmlInputFileName = config["FileSettings:XmlInputFileName"] ?? throw new Exception("XmlInputFileName not configured");

        string inputFilePath = Path.Combine(xmlInputFilePath, xmlInputFileName);
        string outputFolder = xmlOutputFilePath;

        // New file name
        string originalFileName = Path.GetFileNameWithoutExtension(inputFilePath);
        string currentDate = DateTime.Now.ToString("ddMMyyyy");
        string newFileName = $"{originalFileName}_{currentDate}.xls";
        string outputFilePath = Path.Combine(outputFolder, newFileName);

        string mappingFilePath = Path.Combine(assetsFolder, "mappings.json");

        Log.Information("Input file: {InputFile}", inputFilePath);
        Log.Information("Output file: {OutputFile}", outputFilePath);

        // Create output folder if not exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        if (!File.Exists(mappingFilePath))
        {
            Log.Error($"Input file not found: {mappingFilePath}");
            return;
        }

        try
        {
            // Read mappings.json
            string jsonContent = File.ReadAllText(mappingFilePath);

            Dictionary<string, string>? mappings =
                JsonSerializer.Deserialize<Dictionary<string, string>>(jsonContent);

            if (mappings == null)
            {
                Log.Error("Invalid mappings.json");
                return;
            }

            // Open XLS file
            using FileStream file = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read);

            HSSFWorkbook workbook = new HSSFWorkbook(file);

            // First sheet
            ISheet sheet = workbook.GetSheetAt(0);

            // Header row
            IRow headerRow = sheet.GetRow(0);

            // Loop through all columns
            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                ICell cell = headerRow.GetCell(i);

                if (cell == null)
                    continue;

                string currentHeader = cell.StringCellValue.Trim();

                // Check if mapping exists
                if (mappings.ContainsKey(currentHeader))
                {
                    cell.SetCellValue(mappings[currentHeader]);
                }
            }

            // Save updated file
            using FileStream outputFile =
                new FileStream(outputFilePath, FileMode.Create, FileAccess.Write);

            workbook.Write(outputFile);
            Log.Information("File saved successfully");
        }
        catch (Exception ex)
        {
            Log.Error("Error:");
            Log.Error(ex.Message);
        }
        finally
        {
            Log.CloseAndFlush();
        }
    }
}