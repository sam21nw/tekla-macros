using Tekla.Structures.Model;
using Tekla.Structures;

namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
           bool imperial = XSImperialCheck();

		      if (imperial)
		      {
		        Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_IMPERIAL", false);
		      }
		      else
		      {
		        Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_IMPERIAL", true);
		      }
        }
				
				public static bool XSImperialCheck()
    		{
		      bool output = false;
		      string xsImperial = string.Empty;
		      TeklaStructuresSettings.GetAdvancedOption("XS_IMPERIAL", ref xsImperial);
		      if (xsImperial == "TRUE")
		      {
		        output = true;
		      }
		      return output;
    		}
    }
}