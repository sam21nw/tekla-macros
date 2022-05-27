using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;
using Tekla.Structures.Model.UI;
using Tekla.Structures.Geometry3d;
using System.Linq;
using Tekla.BIM.Quantities;

[assembly: Tekla.Technology.Scripting.Compiler.ReferenceAttribute("Tekla.Application.Library")]
//[assembly: Tekla.Technology.Scripting.Compiler.ReferenceAttribute("Tekla.Structures.Datatype")]
//[assembly: Tekla.Technology.Scripting.Compiler.Reference("System.Xml")]
[assembly: Tekla.Technology.Scripting.Compiler.ReferenceAttribute("Tekla.BIM.Toolkit")]

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
    //User prompt string for picking points
    private const string PickPrompt = "Pick all parts to label opening...";
    private const string PickLeaderPosition = "Pick point to place dimension.";

    /// <summary> Color for Dimension lines in model </summary>
    private static readonly Color DimColor = new Color(0.21, 0.82, 0.98);

    /// <summary> Color for Dimension text in model </summary>
    //private static readonly Color TextColor = new Color(0.22, 0.99, 0.61);
    private static readonly Color TextColor = new Color(0.92, 0.11, 0.16); // red color
    //private static readonly Color TextColor = new Color(0.59, 0.96, 0.74); // greenish cyan

		private static readonly Model _model;
		
		public static void RunMacro(IScript akit)
		{
			var pickedObjects = new Picker().PickObjects(Picker.PickObjectsEnum.PICK_N_PARTS, PickPrompt);
			if (pickedObjects.GetSize() < 1) return;

			while (pickedObjects.MoveNext())
			{
				var part = pickedObjects.Current as Part;

				if(part == null)
					continue;

				if(part.GetPartMark().Contains("OPNG")){
					var origin = part.GetCoordinateSystem().Origin;
					var axisX = part.GetCoordinateSystem().AxisX;
					var normal = axisX.GetNormal();
					var axlength = axisX.GetLength();

					double height = 0;
					part.GetReportProperty("HEIGHT", ref height);
					
					double length = 0;
					part.GetReportProperty("LENGTH", ref length);
					
					double width = 0;
					part.GetReportProperty("WIDTH", ref width);
					
					string profile = string.Empty;
					part.GetReportProperty("PROFILE", ref profile);
					
					var profileStr = string.Empty;
					var heightInMM = new Length(height).ToMetricUnits();
					var heightInFtIn = new Length(height).ToImperialUnits();
					var widthInMM = new Length(width).ToMetricUnits();
					var widthInFtIn = new Length(width).ToImperialUnits();

					if (_model.IsImperial())
					{
						if(profile.Contains("D"))
						{
							profileStr = "D" + heightInFtIn;
						}
						else {
							profileStr = heightInFtIn + "X" + widthInFtIn;
						}
					}
					else{
						if(profile.Contains("D"))
						{
							profileStr = "D" + heightInMM;
						}
						else
						{
							profileStr = heightInMM + "X" + widthInMM;
						}
					}

					var midPoint = new Point(
						origin.X + normal.X * axlength /2,
						origin.Y + normal.Y * axlength / 2,
						origin.Z + normal.Z * axlength / 2
						);

					var markPoint = new Point(midPoint);
					markPoint.Z += height / 2 + 250;

					var profilePoint = new Point(markPoint);
					profilePoint.Z -= 100;

					var levelPoint = new Point(markPoint);
					levelPoint.Z -= 200;

					var partMark = part.GetPartMark();

					var drawer = new GraphicsDrawer();

					//drawer.DrawText(markPoint, partMark, TextColor);
					//drawer.DrawText(profilePoint, part.Profile.ProfileString, TextColor);
					drawer.DrawText(profilePoint, profileStr, TextColor);
					//drawer.DrawText(levelPoint, topLevel, DimColor);

					drawer.DrawLineSegment(midPoint, markPoint, DimColor);

				}
			}
		}
	}

	public static class MacroExtensions
  {
    /// <summary>
    /// Gets new point translated from current
    /// </summary>
    /// <param name="pt">Point of origin</param>
    /// <param name="v">Vector translation to move copy of point</param>
    /// <returns>New translated point</returns>
    public static Point GetTranslated(this Point pt, Vector v)
    {
      if (pt == null) throw new ArgumentNullException("pt");
      if (v == null) throw new ArgumentNullException("v");
      var tempPt = new Point(pt);
      tempPt.Translate(v.X, v.Y, v.Z);
      return tempPt;
    }

    /// <summary>
    /// Paints temporary line in the model, Debug model only
    /// </summary>
    /// <param name="ls">Line segment to paint</param>
    /// <param name="col">Color to use, red if skipped</param>
    public static void PaintLine(this LineSegment ls, Color col = null)
    {
      if (col == null) col = new Color(1, 0, 0);
      var gd = new GraphicsDrawer();
      gd.DrawLineSegment(ls.Point1, ls.Point2, col);
    }

    /// <summary>
    /// Translates point vector direcion and distance
    /// </summary>
    /// <param name="pt"></param>
    /// <param name="tVector"></param>
    public static void Translate(this Point pt, Vector tVector)
    {
      pt.Translate(tVector.X, tVector.Y, tVector.Z);
    }

    /// <summary>
    /// Check if is the same direction to other vector
    /// </summary>
    /// <param name="v1">Vector 1</param>
    /// <param name="v2">Vector 2</param>
    /// <param name="angleTolerance">Angle tolerance (radians)</param>
    /// <returns>True if angle is close to 0</returns>
    public static bool IsSameDirection(this Vector v1, Vector v2, double angleTolerance = GeometryConstants.ANGULAR_EPSILON)
    {
      if (v2 == null) return false;
      var angle = v1.GetAngleBetween(v2) * 180 / Math.PI;
      return Math.Abs(angle) < angleTolerance;
    }

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

    /// <summary>
    /// Draws text in the model view
    /// </summary>
    /// <param name="model"></param>
    /// <param name="text">Text to draw</param>
    /// <param name="tPoint">Point to insert as base of text</param>
    /// <param name="tColor">Color to make text</param>
    public static void DrawText(this Model model, string text, Point tPoint, Color tColor)
    {
      DrawText(text, tPoint, tColor);
    }

    /// <summary>
    /// Draws text in the model view
    /// </summary>
    /// <param name="text">Text to draw</param>
    /// <param name="tPoint">Point to insert as base of text</param>
    /// <param name="tColor">Color to make text</param>
    private static void DrawText(string text, Point tPoint, Color tColor)
    {
        var gd = new GraphicsDrawer();
        gd.DrawText(tPoint, text, tColor);
    }

    /// <summary>
    /// Gets midpoint of a line segment
    /// </summary>
    /// <param name="ls">Tekla line segment</param>
    /// <returns>New 3d point at midpoint</returns>
    public static Point GetMidPoint(this LineSegment ls)
    {
        if (ls == null) throw new ApplicationException();
        var startPoint = new Point(ls.Point1);
        var displacement = ls.GetDirectionVector().GetNormal() * ls.Length() * 0.5;
        startPoint.Translate(displacement.X, displacement.Y, displacement.Z);
        return startPoint;
    }
  }

	/// <summary>
  /// Extensions class for Tekla.BIM.Quantities class
  /// </summary>
  public static class TsLength
  {
    /// <summary>
    /// Returns ft-fraction inch string rounded to 1/16 if XS_IMPERIAL=TRUE, mm to decimal places otherwise
    /// </summary>
    public static string ToCurrentUnits(this Length ln)
    {
        return new Model().IsImperial() ? ln.ToString(LengthUnit.Foot, "1/16") : ln.ToString(LengthUnit.Millimeter, "0.00");
    }

    /// <summary>
    /// Uses BIM Length.Parse to convert string to Length and returns Millimeter value
    /// Uses XS_IMPERIAL to decide if input is Foot versus millimeter value in string format
    /// May not work if input is other than foot or mm and flag not set
    /// </summary>
    /// <param name="str">Numberic value in string format</param>
    /// <returns>Double millimeter value</returns>
    public static double FromCurrrentUnits(this string str)
    {
        var length = Length.Parse(str, new Model().IsImperial() ? LengthUnit.Inch : LengthUnit.Millimeter);
        return length.Millimeters;
    }
		
		/// <summary>
    /// Returns ft-fraction inch string rounded to 1/16 in
    /// </summary>
    public static string ToImperialUnits(this Length ln)
    {
        return ln.ToString(LengthUnit.Foot, "1/16");
    }
		
		/// <summary>
    /// Returns mm to decimal places string
    /// </summary>
    public static string ToMetricUnits(this Length ln)
    {
        return ln.ToString(LengthUnit.Millimeter, "0");
    }
  }
}
