using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Email;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Tools.Search;
using OfficeOpenXml;

namespace PSTTest
{
    class Program
    {
        static int totalItems = 0;
        static List<MessageInfo> outputData = new List<MessageInfo>();

        static void Main(string[] args)
        {
            string pstPath = "D:\\MIP\\TestPST.PST";
            var pst = Aspose.Email.Storage.Pst.PersonalStorage.FromFile(pstPath);
            var rootFolder = pst.RootFolder;

            CountContents(rootFolder);
            WriteExcel(outputData);
            Console.WriteLine(totalItems);
            Console.ReadKey();
        }

        static void CountContents(FolderInfo folder)
        {
            
            FolderInfoCollection rootFolders = folder.GetSubFolders();
            totalItems += folder.ContentCount;

            if (folder.ContentCount > 0)
            {
                for (int i = 0; i < folder.ContentCount; i += 50)
                {
                    foreach (var message in folder.GetContents(i, 50))
                    {
                        var mi = new MessageInfo { Id = message.EntryIdString, Subject = message.Subject };
                        outputData.Add(mi);
                    }
                }
            }

            Console.WriteLine(folder.DisplayName + ": "+ folder.ContentCount.ToString());

            foreach (FolderInfo subFolder in rootFolders)
            {
                // totalItems += subFolder.GetSubFolders().Count();
                CountContents(subFolder);
            }
        }
        static void WriteExcel(List<MessageInfo> data)
        {
            var file = new System.IO.FileInfo(".\\output.xlsx");
            if(file.Exists) { file.Delete(); }

            using (var Excel = new ExcelPackage(file))
            {
                var Worksheet = Excel.Workbook.Worksheets.Add("MyData");
                Worksheet.Cells["A1"].LoadFromCollection(data, true, OfficeOpenXml.Table.TableStyles.Dark10);
                Excel.Save();
            }            
        }
    }
    public class MessageInfo
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public int AttachmentCount { get; set; }
    }
}
