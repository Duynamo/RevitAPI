
#region Namespaces

using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Application = Autodesk.Revit.ApplicationServices.Application;

#endregion

namespace AlphaBIM
{
    [Transaction(TransactionMode.Manual)]
    public class SplitPipeCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
            ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument    uidoc = uiapp.ActiveUIDocument;
            Application   app   = uiapp.Application;
            Document      doc   = uidoc.Document;

            // Khi chạy bằng Add-in Manager thì comment 2 dòng bên dưới để tránh lỗi
            string dllFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            AssemblyLoader.LoadAllRibbonAssemblies(dllFolder);

            try
            {
                // ExternalEvent handler – thực thi Split/Flange trên Revit main thread
                var handler = new SplitPipeHandler();
                var exEvent = ExternalEvent.Create(handler);

                // Show() → modeless window, Execute() trả về ngay
                // → Revit nhận lại focus → PickObject hoạt động bình thường
                var window = new SplitPipeWindow(uidoc, doc, handler, exEvent);
                window.Show();
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }
}
