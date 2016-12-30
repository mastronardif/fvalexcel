using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.File;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebJob12
{
    class ValidateExcel
    {

        public static void fileShareExample()
        {
            string connectionString = AmbientConnectionStringProvider.Instance.GetConnectionString(ConnectionStringNames.Storage);
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);

            // Create a CloudFileClient object for credentialed access to File storage.
            CloudFileClient fileClient = storageAccount.CreateCloudFileClient();

            // Get a reference to the file share we created previously.
            CloudFileShare share = fileClient.GetShareReference("myfileshare");

            String azFn = "testcampaign.xlsx";
            // Ensure that the share exists.
            if (share.Exists())
            {
                CloudFileDirectory rootDir = share.GetRootDirectoryReference();
                CloudFileDirectory sampleDir = rootDir.GetDirectoryReference(".");
                CloudFile sourceFile = sampleDir.GetFileReference(azFn); // "bobotheclown.txt");

                // Ensure that the source file exists.
                if (sourceFile.Exists())
                {
                    // Get a reference to the destination file.
                    //CloudFile destFile = sampleDir.GetFileReference("Log1Copy.txt");

                    // Start the copy operation.
                    //string data = sourceFile.DownloadText(); //BeginDownloadText(sourceFile, );
                    //string output = @"c:\fxm\downloads\az_" + azFn;
                    //sourceFile.DownloadToFile(output, FileMode.Create);
                    //FileStream fs11 = new FileStream(;


                    // Write the contents of the destination file to the console window.
                    //Console.WriteLine(sourceFile.DownloadText());
                    MemoryStream mem = new MemoryStream();
                    sourceFile.DownloadToStream(mem);
                    mem.Position = 0;
                    //FileStream fs = new FileStream(output, FileMode.Open);

                    //List<string> ColNames = new List<string>(){ "Campaigns", "Impressions", "Clicks" };
                    List<string> ColNames = new List<string>() { "Campaigns", "Impressions", "Clicks" };
                    //testExcel(fs, ColNames);
                    testExcel(mem, ColNames);
                }
            }
        }

        //private static void testExcel(FileStream fileStream, List<string> ColNames)
        private static void testExcel(MemoryStream fileStream, List<string> ColNames)
        {
            List<string> header = new List<string>();

            using (var package = new ExcelPackage(fileStream))
            {
                // Get the workbook in the file
                var workbook = package.Workbook;

                //workbook.Worksheets[1];
                //workbook.Worksheets[1].
                //var ws = package.Workbook.Worksheets["Worksheet1"];
                var ws = package.Workbook.Worksheets[1];
                //DataTable tbl = new DataTable();
                var hasHeader = true;
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    //tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                    header.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }

                Console.WriteLine(String.Join(", ", header));
                //bool hasMatch = ContainsAllItems(header, ColNames);
                //bool hasMatch = !header.Except(ColNames).Any();
                bool containsAll = ColNames.All(s => header.Contains(s));
                //bool hasMatch = header.Any(x => ColNames.Any(y => y == x));
                Console.WriteLine(string.Format("containsAll({0})", containsAll));

                //var startRow = hasHeader ? 2 : 1;
                //for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                //{
                //    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                //    var row = tbl.NewRow();
                //    foreach (var cell in wsRow)
                //    {
                //        row[cell.Start.Column - 1] = cell.Text;
                //    }
                //    tbl.Rows.Add(row);
                //}
            }
        }


    }
}
