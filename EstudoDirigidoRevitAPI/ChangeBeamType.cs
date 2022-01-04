using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EstudoDirigidoRevitAPI
{
    [Transaction(TransactionMode.Manual)]
    public class ChangeBeamType : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<Element> beams = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_StructuralFraming).ToElements().ToList();

            FamilySymbol familySymbol = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .Where(fs => fs.Name == "VA 13x30cm")
                .First() as FamilySymbol;

            List<ElementId> beamsIds = beams.Select(beam => beam.Id).ToList();

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("Alterar tipo das vigas");

                Element.ChangeTypeId(doc, beamsIds, familySymbol.Id);

                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}
