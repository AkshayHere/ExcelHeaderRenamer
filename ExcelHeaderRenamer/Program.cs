using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text.Json;

class Program
{
    static void Main(string[] args)
    {
        // Project root
        string projectRoot = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory)!
            .Parent!
            .Parent!
            .Parent!
            .FullName;

        // Assets folder
        string assetsFolder = Path.Combine(projectRoot, "Assets");

        string inputFilePath = @"C:\Projects\EXTZRRPF2.xls";
        string outputFolder = @"C:\Projects\Save";

        // New file name
        string originalFileName = Path.GetFileNameWithoutExtension(inputFilePath);
        string currentDate = DateTime.Now.ToString("ddMMyyyy");
        string newFileName = $"{originalFileName}_{currentDate}.xls";
        string outputFilePath = Path.Combine(outputFolder, newFileName);

        string mappingFilePath = Path.Combine(assetsFolder, "mappings.json");

        // Create output folder if not exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        if (!File.Exists(mappingFilePath))
        {
            Console.WriteLine($"Input file not found: {mappingFilePath}");
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
                Console.WriteLine("Invalid mappings.json");
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

            Console.WriteLine("Headers updated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error:");
            Console.WriteLine(ex.Message);
        }
    }
}