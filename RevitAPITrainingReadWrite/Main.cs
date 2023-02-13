using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
using File = System.IO.File;

namespace RevitAPITrainingReadWrite
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;



            OpenFileDialog openFileDialog1 = new OpenFileDialog 
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "Excel files(*.xlsx) | *.xlsx" 
            };

            string filePath =string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;
            }

            if (string.IsNullOrEmpty(filePath))
                return Result.Cancelled;

            var rooms =new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Rooms)
            .Cast<Room>()
            .ToList();

            using(FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(filePath); // создаём переменную для чтения файла экселя по указанному пути
                ISheet sheet = workbook.GetSheetAt(index:0); // берём лист из которого будем читать данные. В данном случае - самый первый в файле.

                // построчно проходим по листу и собираем данные
                int rowIndex = 0;
                while (sheet.GetRow(rowIndex)!=null)
                {
                    if (sheet.GetRow(rowIndex).GetCell(0) == null ||
                        sheet.GetRow(rowIndex).GetCell(1) == null)
                    {
                        rowIndex++;
                        continue;
                    }
                    string name = sheet.GetRow(rowIndex).GetCell(0).StringCellValue;
                    string number= sheet.GetRow(rowIndex).GetCell(1).StringCellValue;

                    var room = rooms.FirstOrDefault(r => r.Equals(number));// берём первое помещение из модели, у которого номер совпадает с тем номером,
                                                                           // который находится в текущей строке

                    if (room == null)
                    {
                        rowIndex++;
                        continue;
                    }

                    using (var ts = new Transaction(doc, "Set parameter"))
                    {
                        ts.Start();
                        room.get_Parameter(BuiltInParameter.ROOM_NUMBER).Set(name); 
                        ts.Commit();

                    }

                    rowIndex++;
                }

            }

            return Result.Succeeded;
        }
    }
}
