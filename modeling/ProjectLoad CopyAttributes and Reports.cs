using System.IO;
using Tekla.Structures.Model;
using Tekla.Structures;

namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
				 CopyAttributes();
		    }
				
				public static void CopyAttributes(){
	      Model model = new Model();
	      ModelInfo info = model.GetInfo();
	      var modelPath = info.ModelPath;
				
				string attrSourcePath = "Z:\\Grating Users\\SAM\\Settings\\Tekla\\02_OMM_Tekla_Settings\\01_attributes";
				string reportsSourcePath = "Z:\\Grating Users\\SAM\\Settings\\Tekla\\02_OMM_Tekla_Settings\\04_reports and templates";
				
	      string ModelAttributesPath = Path.Combine(modelPath, "attributes");
				if (!Directory.Exists(ModelAttributesPath)) Directory.CreateDirectory(ModelAttributesPath);
				
				string fileName = string.Empty;
				string destFile = string.Empty;
				
				try
				{
					if (Directory.Exists(attrSourcePath))
					{
				    string[] files = Directory.GetFiles(attrSourcePath);
				    // Copy the files and overwrite destination files if they already exist. 
				    foreach (string s in files)
				    {
			        // Use static Path methods to extract only the file name from the path.
			        fileName = Path.GetFileName(s);
			        destFile = Path.Combine(ModelAttributesPath, fileName);
			        File.Copy(s, destFile, true);
				    }
					}
					if (Directory.Exists(reportsSourcePath))
					{
				    string[] files = Directory.GetFiles(reportsSourcePath);
				    // Copy the files and overwrite destination files if they already exist. 
				    foreach (string s in files)
				    {
			        // Use static Path methods to extract only the file name from the path.
			        fileName = Path.GetFileName(s);
			        destFile = Path.Combine(modelPath, fileName);
			        File.Copy(s, destFile, true);
				    }
					}
				}
				catch
				{
					
				}
  		}
		}
}