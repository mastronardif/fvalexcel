using OfficeOpenXml;
public static void Run(Stream myBlob, string name, TraceWriter log)
{
    log.Info($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");
    log.Info($"C# Blob trigger function processed: {myBlob}");

    MemoryStream mem = new MemoryStream();
    myBlob.CopyTo(mem);
    mem.Position = 0;
    List<string> ColNames = new List<string>() { "Campaigns", "Impressions", "Clicks" };
    ValidateExcel.testExcel(log, mem, ColNames);
}

class ValidateExcel
{
    public static void testExcel(TraceWriter log, MemoryStream fileStream, List<string> ColNames)
    {
        List<string> header = new List<string>();

        using (var package = new ExcelPackage(fileStream))
        {
            // Get the workbook in the file
            var workbook = package.Workbook;

            var ws = package.Workbook.Worksheets[1];
            var hasHeader = true;
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
            {
                header.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }

            log.Info(String.Join(", ", header));
            bool containsAll = ColNames.All(s => header.Contains(s));
            log.Info(string.Format("containsAll({0})", containsAll));
        }
    }
}