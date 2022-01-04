using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EstudoDirigidoRevitAPI
{
    [Transaction(TransactionMode.Manual)]
    public class CollectBeamData : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<Element> beams = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_StructuralFraming).ToElements().ToList();

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "Excel file|*.xlsx";

            saveFileDialog.Title = "Escolha o local para salvar o arquivo";

            saveFileDialog.ShowDialog();

            using (StreamWriter sw = new StreamWriter(saveFileDialog.OpenFile()))
            {
                foreach (Element beam in beams)
                {
                    string floor = beam.LookupParameter("Pavimento (Mat)").AsString();
                    double length = Math.Round(UnitUtils.ConvertFromInternalUnits(beam.LookupParameter("Comprimento").AsDouble(), UnitTypeId.Meters), 2);

                    string line = $"{beam.Id};{floor};{length}";
                    sw.WriteLine(line);
                }
            }

            return Result.Succeeded;
        }
    }
}
