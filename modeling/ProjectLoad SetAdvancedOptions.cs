using Tekla.Structures.Model;
using Tekla.Structures;

namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
		        SetAdvancedOptions();
        }
				
				public static void SetAdvancedOptions()
				{
					Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_CHECK_FLAT_LENGTH_ALSO", false);
					Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_USE_FLAT_DESIGNATION", false);
					Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_USE_NEW_PLATE_DESIGNATION", "");
					Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_UNIQUE_ASSEMBLY_NUMBERS", false);
					Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_UNIQUE_NUMBERS", false);
					Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_USE_ASSEMBLY_NUMBER_FOR", "MAIN_PART");
					//Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_FIRM", "");
					Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_ASSEMBLY_POSITION_NUMBER_FORMAT_STRING", 					"%ASSEMBLY_PREFIX%%ASSEMBLY_POS.3%");
					Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_ALLOW_INCH_MARK_IN_DIMENSIONS", true);
					Tekla.Structures.ModelInternal.Operation.dotSetAdvancedOption("XS_INCH_SIGN_ALWAYS", true);
				}
    }
}
