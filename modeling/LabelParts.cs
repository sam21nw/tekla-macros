using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using Tekla.Structures;
//using TD = Tekla.Structures.Datatype;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;
using Tekla.Structures.Model.UI;
using Tekla.Structures.Geometry3d;
//using System.Globalization;
//using System.Xml;
//using System.Windows.Forms;
//using System.Linq;
//using Tekla.BIM.Quantities;

[assembly: Tekla.Technology.Scripting.Compiler.ReferenceAttribute("Tekla.Application.Library")]
//[assembly: Tekla.Technology.Scripting.Compiler.ReferenceAttribute("Tekla.Structures.Datatype")]
//[assembly: Tekla.Technology.Scripting.Compiler.Reference("System.Xml")]
//[assembly: Tekla.Technology.Scripting.Compiler.ReferenceAttribute("Tekla.BIM.Toolkit")]

namespace Tekla.Technology.Akit.UserScript
{
  public class Script
  {
    public static void Run(Tekla.Technology.Akit.IScript akit)
    {
			try
			{
				LabelParts.RunMacro(akit);
			}
			catch (Exception ex)
			{
				Trace.WriteLine(ex.InnerException + ex.Message + ex.StackTrace);
			}
    }
  }
	
	public static class LabelParts
	{
    /// <summary> Color for Dimension lines in model </summary>
    private static readonly Color DimColor = new Color(0.21, 0.82, 0.98);

    /// <summary> Color for Dimension text in model </summary>
    private static readonly Color TextColor = new Color(0.59, 0.96, 0.74);
		
		public static void RunMacro(IScript akit)
		{
			var pickedObjects = new Picker().PickObjects(Picker.PickObjectsEnum.PICK_N_PARTS);
			if (pickedObjects.GetSize() < 1) return;

			while (pickedObjects.MoveNext())
			{
				var part = pickedObjects.Current as Part;

				if(part == null)
					continue;

				var origin = part.GetCoordinateSystem().Origin;
				var axisX = part.GetCoordinateSystem().AxisX;
				var normal = axisX.GetNormal();
				var length = axisX.GetLength();

				double height = 0;
				part.GetReportProperty("HEIGHT", ref height);
				
				string name = string.Empty;
				part.GetReportProperty("NAME", ref name);

				string topLevel = string.Empty;
				part.GetReportProperty("TOP_LEVEL", ref topLevel);
				
				string profile = string.Empty;
				part.GetReportProperty("PROFILE", ref profile);

				var midPoint = new Point(
					origin.X + normal.X * length /2,
					origin.Y + normal.Y * length / 2,
					origin.Z + normal.Z * length / 2
					);

				var markPoint = new Point(midPoint);
				markPoint.Z += height / 2 + 250;

				var profilePoint = new Point(markPoint);
				profilePoint.Z -= 100;

				var levelPoint = new Point(markPoint);
				levelPoint.Z -= 200;

				var partMark = part.GetPartMark();

				var drawer = new GraphicsDrawer();

				//drawer.DrawText(markPoint, name, TextColor);
				//drawer.DrawText(profilePoint, part.Profile.ProfileString, TextColor);
				drawer.DrawText(profilePoint, profile, TextColor);
				//drawer.DrawText(levelPoint, topLevel, DimColor);
				drawer.DrawLineSegment(midPoint, markPoint, DimColor);
			}
		}
	}
	
	public static class MacroExtensions
	{
		 /// <summary>
    /// Returns if imperial units are being used
    /// </summary>
    public static bool IsImperial(this Model model)
    {
      var stringTemp = string.Empty;
      TeklaStructuresSettings.GetAdvancedOption("XS_IMPERIAL", ref stringTemp);
      if (!string.IsNullOrEmpty(stringTemp)) return true;
      return string.CompareOrdinal(stringTemp, "1") == 0;
    }
	}
}
