using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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

            string roomInfo=string.Empty;
            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();
            foreach (Room room in rooms)
            {
                string roomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                roomInfo += $"{roomName}\t{room.Number}\t{room.Area}{Environment.NewLine}";
            }

            var saveFileDialog = new SaveFileDialog
            {
                OverwritePrompt = true,//если файл существует, выдавать запрос на его перезапись
                InitialDirectory=Environment.GetFolderPath(Environment.SpecialFolder.Desktop),//начальная папка, скоторой будет начинаться диалог
                Filter = "All files(*.*)|*.*",//фильтр отображаемых файлов по расширению
                FileName="roomInfo.csv",//Имя файла по умолчанию
                DefaultExt=".csv"//Расширение по умолчани
            };
            string selectedFilePatch=string.Empty;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePatch = saveFileDialog.FileName;
            }
            if (string.IsNullOrEmpty(selectedFilePatch))
                return Result.Cancelled;

            File.WriteAllText(selectedFilePatch, roomInfo);

            return Result.Succeeded;
        }
    }
}
