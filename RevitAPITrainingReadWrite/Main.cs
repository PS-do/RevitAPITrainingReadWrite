﻿using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
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


            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "Json files (*.json) | *.json"
            };

            string filePath = string.Empty;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
            }

            if (string.IsNullOrEmpty(filePath))
            {
                return Result.Cancelled;
            }

            string json = File.ReadAllText(filePath);
            List<RoomData> roomDataList = JsonConvert.DeserializeObject<List<RoomData>>(json);

            var rooms = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Rooms)
            .Cast<Room>()
            .ToList();

            foreach (RoomData roomData in roomDataList)
            {
                var room = rooms.FirstOrDefault(r => r.Number.Equals(roomData.Number));
                if (room == null)
                    continue;

                using (var ts = new Transaction(doc, "Set parameter"))
                {
                    ts.Start();
                    room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(roomData.Name);
                    ts.Commit();
                }
            }



            return Result.Succeeded;
        }
    }
}
