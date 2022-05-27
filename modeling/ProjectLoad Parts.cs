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
				
				string partsSourcePath = "Z:\\Grating Users\\SAM\\Settings\\Tekla\\02_OMM_Tekla_Settings\\03_parts";
				
	      string ModelAttributesPath = Path.Combine(modelPath, "attributes");
				if (!Directory.Exists(ModelAttributesPath)) Directory.CreateDirectory(ModelAttributesPath);
				
				string fileName = string.Empty;
				string destFile = string.Empty;
				
				try
				{
					if (Directory.Exists(partsSourcePath))
					{
					    string[] files = Directory.GetFiles(partsSourcePath);
					    // Copy the files and overwrite destination files if they already exist. 
					    foreach (string s in files)
					    {
				        // Use static Path methods to extract only the file name from the path.
				        fileName = Path.GetFileName(s);
				        destFile = Path.Combine(ModelAttributesPath, fileName);
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