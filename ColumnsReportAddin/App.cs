using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace ColumnsReportAddin
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            string tabName = "KAITECH-BD-R10";
            app.CreateRibbonTab(tabName);

            RibbonPanel panel = app.CreateRibbonPanel(tabName, "Structure");

            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData(
                "ColumnsReport",
                "Columns\nReport",
                assemblyPath,
                "ColumnsReportAddin.ColumnsExporter");

            PushButton button = panel.AddItem(buttonData) as PushButton;
            button.ToolTip = "Export Columns Report to Excel";
            button.LargeImage = LoadEmbeddedImage("ColumnsReportAddin.Resources.column.png");

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }

        private BitmapImage LoadEmbeddedImage(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;
                BitmapImage bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.StreamSource = stream;
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.EndInit();
                return bmp;
            }
        }
    }
}
