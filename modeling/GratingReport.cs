using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Tekla.Structures;
using TSM = Tekla.Structures.Model;
using Tekla.Structures.Model;

namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            var CurrentModel = new TSM.Model();
            var dataDir = "";
            TeklaStructuresSettings.GetAdvancedOption("XSDATADIR", ref dataDir);
            if (dataDir == null || string.IsNullOrEmpty(dataDir))
            {
                MessageBox.Show("Unable to resolve the XSDATADIR advanced option - unable to run the application.");
                return;
            }

            const string ApplicationName = "GratingReport.exe";
            var applicationPath = Path.Combine(dataDir, "Environments\\common\\extensions\\OASIS", ApplicationName);
                
			//MessageBox.Show(applicationPath);

            if (File.Exists(applicationPath))
            {
                var newProcess = new Process { StartInfo = { FileName = applicationPath } };

                try
                {
                    newProcess.Start();
                }
                catch
                {
                    MessageBox.Show(ApplicationName + " failed to start.");
                }
            }
            else
            {
                MessageBox.Show(ApplicationName + " not found.");
            }
        }
    }
}