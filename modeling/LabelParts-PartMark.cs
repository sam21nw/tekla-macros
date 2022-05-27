using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;
using Tekla.Structures.Model.UI;
using Tekla.Structures.Geometry3d;

[assembly: Tekla.Technology.Scripting.Compiler.ReferenceAttribute("Tekla.Application.Library")]

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
    private static readonly Color TextColor = new Color(0.92, 0.11, 0.16); // red color
    //private static readonly Color TextColor = new Color(1.00, 1.00, 0.04); // cyan
		
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

				drawer.DrawText(markPoint, partMark, TextColor);
				//drawer.DrawText(profilePoint, profile, TextColor);

				drawer.DrawLineSegment(midPoint, markPoint, DimColor);
			}
		}
	}
}
