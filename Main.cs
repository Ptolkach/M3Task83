using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace M3Task83
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            using (Transaction tx = new Transaction(doc))
            {

                tx.Start("Export Image");

                string desktop_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                ViewPlan viewPlan = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewPlan))
                    .Cast<ViewPlan>()
                    .FirstOrDefault(v => v.ViewType == ViewType.FloorPlan && v.Name.Equals("Level 1"));

                string filepath = Path.Combine(desktop_path, viewPlan.Name);

                ImageExportOptions img = new ImageExportOptions();

                img.ZoomType = ZoomFitType.FitToPage;
                img.PixelSize = 1024;
                img.ImageResolution = ImageResolution.DPI_600;
                img.FitDirection = FitDirectionType.Horizontal;
                img.ExportRange = ExportRange.CurrentView;
                img.HLRandWFViewsFileType = ImageFileType.PNG;
                img.FilePath = filepath;
                img.ShadowViewsFileType = ImageFileType.PNG;

                doc.ExportImage(img);

                tx.RollBack();

                filepath = Path.ChangeExtension(filepath, "png");

                Process.Start(filepath);
                 
            }
            return Result.Succeeded;
        }
    }
}
