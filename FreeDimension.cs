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
using System.Linq;
using Tekla.BIM.Quantities;

[assembly: Tekla.Technology.Scripting.Compiler.ReferenceAttribute("Tekla.Application.Library")]
//[assembly: Tekla.Technology.Scripting.Compiler.ReferenceAttribute("Tekla.Structures.Datatype")]
//[assembly: Tekla.Technology.Scripting.Compiler.Reference("System.Xml")]
[assembly: Tekla.Technology.Scripting.Compiler.ReferenceAttribute("Tekla.BIM.Toolkit")]
namespace Tekla.Technology.Akit.UserScript
{
	/// <summary>
	/// Internal class for running logic
	/// </summary>
	public class Script
	{
		/// <summary>
		/// Internal method run automatically by Tekla Structures if using as raw c# file
		/// </summary>
		/// <param name="akit">Passed argument automatically by core when using as macro</param>
		public static void Run(IScript akit)
		{
		try
			{
				FreeDimension.RunMacro(akit);
			}
			catch (Exception ex)
			{
				Trace.WriteLine(ex.InnerException + ex.Message + ex.StackTrace);
			}
		}
	}

	public static class FreeDimension
	{
    //User prompt string for picking points
    private const string PickPrompt = "Pick points to dimension.";
    private const string PickLeaderPosition = "Pick point to place dimension.";

    /// <summary> Color for Dimension lines in model </summary>
    private static readonly Color DimColor = new Color(0.98, 0.31, 0.16);

    /// <summary> Color for Dimension text in model </summary>
    private static readonly Color TextColor = new Color(1.00, 1.00, 0.04);
				
		public static void RunMacro(IScript akit)
		{
			try
			{
	      //Prompt user for points to dimension
	      var pts = PickerService.PickTwoPoints(PickPrompt);
	      if (pts == null || pts.Points.Count < 2) return;
        var pt1 = pts.Points[0] as Point;
        var pt2 = pts.Points[1] as Point;
				
        //Prompt user for point to place dimension
        var leaderPt = PickerService.PickAPoint(PickLeaderPosition);
        if (leaderPt == null) return;
				
				if (pt1 == null) return;
				if (pt2 == null) return;
				
				var dist = Math.Abs(Distance.PointToPoint(pt2, pt1));
				var distX = Math.Abs(pt2.X-pt1.X);
				var distY = Math.Abs(pt2.Y-pt1.Y);
				var distZ = Math.Abs(pt2.Z-pt1.Z);
				
				var distInMM = new Length(dist).ToMetricUnits();
				var distInMMXdir = new Length(distX).ToMetricUnits();
				var distInMMYdir = new Length(distY).ToMetricUnits();
				var distInMMZdir = new Length(distZ).ToMetricUnits();
				
				string distInFtIn = new Length(dist).ToImperialUnits();
				string distInFtInXdir = new Length(distX).ToImperialUnits();
				string distInFtInYdir = new Length(distY).ToImperialUnits();
				string distInFtInZdir = new Length(distZ).ToImperialUnits();
				
				string ftInString = "ft-in";
				string ftInStringX = "ft-in";
				string ftInStringY = "ft-in";
				string ftInStringZ = "ft-in";
				
				ftInString = (dist < 304.8) ? "in" : "ft-in";
				ftInStringX = (distX < 304.8) ? "in" : "ft-in";
				ftInStringY = (distY < 304.8) ? "in" : "ft-in";
				ftInStringZ = (distZ < 304.8) ? "in" : "ft-in";
				
				string displayDistString = "Distance:-  " + distInMM + " " + "(" + distInFtIn + ")";
				string displayStringX = "[X:" + distInMMXdir + " (" + distInFtInXdir + ")]";
				string displayStringY = "[Y:" + distInMMYdir + " (" + distInFtInYdir + ")]";
				string displayStringZ = "[Z:" + distInMMZdir + " (" + distInFtInZdir + ")]";
				
				//Display distance in model prompt
				Operation.DisplayPrompt(displayDistString + "   " + displayStringX + "  " + displayStringY + "  " + displayStringZ);
				
        var xAxis = new Vector(pt2 - pt1);
        var yAxis = Vector.Cross(new Vector(0, 0, 1), xAxis);
        var wp = new CoordinateSystem(pt1, xAxis, yAxis);
				
        //Call to create dimension graphics in model
        var modelDim = new OrthogonalDimensionSet(pts.Points.Cast<Point>().ToList(), leaderPt, wp, DimColor, TextColor);
        if (modelDim.Create()) return;

        //If failed to create, inform user
        Trace.WriteLine("Unable to create model dimension set.");
        Operation.DisplayPrompt("Unable to create model dimension set.");
			}
				catch (Exception ex)
			{
				//Needed for picker interrupt
				Operation.DisplayPrompt(ex.ToString());
			}
			finally
			{
        TeklaStructures.Disconnect();
			}
		}
	}
	
	/// <summary>
  /// Model graphic dimensioning tool set
  /// </summary>
  public class OrthogonalDimensionSet
  {
    private readonly List<Point> _pts;
    private readonly Color _primaryColor;
    private readonly Color _secondaryColor;
    private readonly Point _leaderPoint;
    private const double DefaultOffsetDistance = 500.0;

    private const double TickLength = 100.0;
    private const double TickAngle = 45.0;
    private const double LeaderExtension = 141.42;

    /// <summary>
    /// New dimension set with property sets filled
    /// </summary>
    /// <param name="pts">Points to dimension</param>
    /// <param name="leaderPt">Point to position dimension leader line</param>
    /// <param name="cs">Dimension coordinate sytem</param>
    /// <param name="dimColor">Color of dimension lines</param>
    /// <param name="txtColor">Color of text</param>
    public OrthogonalDimensionSet(List<Point> pts, Point leaderPt, CoordinateSystem cs, Color dimColor, Color txtColor)
    {
      if (pts == null) throw new ArgumentNullException("pts");
      if (leaderPt == null) throw new ArgumentNullException("leaderPt");
      if (cs == null) throw new ArgumentNullException("cs");
      if (dimColor == null) throw new ArgumentNullException("dimColor");
      if (txtColor == null) throw new ArgumentNullException("txtColor");

      _pts = pts;
      _primaryColor = dimColor;
      _secondaryColor = txtColor;
      _leaderPoint = leaderPt;
    }

    /// <summary>
    /// Draws temporary graphics in model
    /// </summary>
    /// <returns>True if created</returns>
    public bool Create()
    {
      return DrawInModel();
    }

    /// <summary>
    /// Draws temporary graphics in model
    /// </summary>
    /// <returns>True if created</returns>
    private bool DrawInModel()
    {
      if (_pts == null || _pts.Count < 2) return false;
      if (_leaderPoint == null) return false;
      if (_primaryColor == null || _secondaryColor == null) return false;

      Point lastPoint = null;
      var xAxis = new Vector(_pts[1] - _pts[0]).GetNormal();
      var yAxisTemp = new Vector(_leaderPoint - _pts[0]).GetNormal();
      var zAxisTemp = Vector.Cross(xAxis, yAxisTemp);
      var yAxis = Vector.Cross(zAxisTemp, xAxis);
      var offsetDist = _leaderPoint == null ? DefaultOffsetDistance : Distance.PointToLine(_leaderPoint, new Line(_pts[0], _pts[1]));
      var newCs = new CoordinateSystem(new Point(), xAxis, yAxis);

      foreach (var pt in _pts)
      {
        if (lastPoint != null) DrawDimensionInModel(lastPoint, pt, offsetDist, newCs, _primaryColor, _secondaryColor);
        lastPoint = pt;
      }
      return true;
    }

    /// <summary>
    /// Creates orthogonal 2 point dimension in model
    /// </summary>
    /// <param name="p1">Point 1</param>
    /// <param name="p2">Point 2</param>
    /// <param name="offsetDist">Distance from point to line offset</param>
    /// <param name="dimCs">Coordinate system for dimension line (x axis = leader line, y axis = offset line direction)</param>
    /// <param name="dimColor">Dimension color</param>
    /// <param name="txtColor">Text color</param>
    private static void DrawDimensionInModel(Point p1, Point p2, double offsetDist, CoordinateSystem dimCs, Color dimColor, Color txtColor)
    {
      if (p1 == null) throw new ArgumentNullException("p1");
      if (p2 == null) throw new ArgumentNullException("p2");
      if (dimCs == null) throw new ArgumentNullException("dimCs");
      if (dimColor == null) throw new ArgumentNullException("dimColor");
      if (txtColor == null) throw new ArgumentNullException("txtColor");

      //If points are not in order of passed coord sys x dir, then swap x axis
      if (!dimCs.AxisX.IsSameDirection(new Vector(p2 - p1), 5)) dimCs.AxisX *= -1;

      var distance = Distance.PointToPoint(p1, p2);
      var leg1 = new LineSegment(p1, p1.GetTranslated(dimCs.AxisY.GetNormal() * offsetDist));
      var leg2 = new LineSegment(leg1.Point2, leg1.Point2.GetTranslated(dimCs.AxisX.GetNormal() * distance));
      var leg3 = new LineSegment(leg2.Point2, p2);
      leg1.PaintLine(dimColor);
      leg2.PaintLine(dimColor);
      leg3.PaintLine(dimColor);

      //Get angle dimensions for tick marks
      var dx = Math.Cos(TickLength * Math.PI / 180) * TickLength * 0.5;
      var dy = Math.Sin(TickAngle * Math.PI / 180) * TickLength * 0.5;

      var tick1 = new LineSegment(new Point(leg1.Point2), new Point(leg1.Point2));
      tick1.Point1.Translate(leg2.GetDirectionVector() * -dx);
      tick1.Point1.Translate(leg1.GetDirectionVector() * -dy);
      tick1.Point2.Translate(leg2.GetDirectionVector() * dx);
      tick1.Point2.Translate(leg1.GetDirectionVector() * dy);
      tick1.PaintLine(dimColor);

      var tick2 = new LineSegment(new Point(leg2.Point2), new Point(leg2.Point2));
      tick2.Point1.Translate(leg2.GetDirectionVector() * -dx);
      tick2.Point1.Translate(leg3.GetDirectionVector() * dy);
      tick2.Point2.Translate(leg2.GetDirectionVector() * dx);
      tick2.Point2.Translate(leg3.GetDirectionVector() * -dy);
      tick2.PaintLine(dimColor);

      var startEndExtHoriz = new LineSegment(new Point(leg1.Point2), new Point(leg1.Point2));
      startEndExtHoriz.Point1.Translate(leg2.GetDirectionVector() * -LeaderExtension);
      startEndExtHoriz.PaintLine(dimColor);

      var finEndExtHoriz = new LineSegment(new Point(leg2.Point2), new Point(leg2.Point2));
      finEndExtHoriz.Point2.Translate(leg2.GetDirectionVector() * LeaderExtension);
      finEndExtHoriz.PaintLine(dimColor);

      var startEndExtVert = new LineSegment(new Point(leg1.Point2), new Point(leg1.Point2));
      startEndExtVert.Point1.Translate(leg1.GetDirectionVector() * LeaderExtension);
      startEndExtVert.PaintLine(dimColor);

      var finEndExtVert = new LineSegment(new Point(leg2.Point2), new Point(leg2.Point2));
      finEndExtVert.Point2.Translate(leg3.GetDirectionVector() * -LeaderExtension);
      finEndExtVert.PaintLine(dimColor);

      MacroExtensions.DrawText(new Model(), new Length(distance).ToCurrentUnits(), leg2.GetMidPoint(), txtColor);
    }

    /// <summary>
    /// Draws line segment in model with temporary graphics
    /// </summary>
    /// <param name="ls">Line segment</param>
    /// <param name="color">Color to use</param>
    /// <returns>True if successful</returns>
    protected static bool DrawLineSegment(LineSegment ls, Color color)
    {
      if (ls == null) throw new ArgumentNullException("ls");
      if (color == null) throw new ArgumentNullException("color");
      var gd = new GraphicsDrawer();
      return gd.DrawLineSegment(ls.Point1, ls.Point2, color);
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
	
  public static class PickerService
  {
    /// <summary>Tekla Structures Model Internal Picker </summary>
    public static readonly Picker ModelPicker = new Picker();

    /// <summary>
    /// Picks a polygon
    /// </summary>
    /// <param name="prompt">
    /// The prompt
    /// </param>
    /// <returns>
    /// A Polygon of the picked points
    /// </returns>
    public static Polygon PickAPolygon(string prompt = null)
    {
      ArrayList pick = null;
      Polygon polygon = null;
      try
      {
        while (pick == null)
        {
          if (prompt == null) pick = ModelPicker.PickPoints(Picker.PickPointEnum.PICK_POLYGON);
          else pick = ModelPicker.PickPoints(Picker.PickPointEnum.PICK_POLYGON, prompt);
        }

        if (pick.Count >= 2)
        {
          polygon = new Polygon();
          polygon.Points.AddRange(pick);
        }
      }
      catch (Exception ex)
      {
        if (ex.Message.Contains("interrupt")) { }
        else Trace.WriteLine(ex.Message);
      }
      return polygon;
    }

    /// <summary>
    /// Picks a point
    /// </summary>
    /// <param name="prompt">
    /// The prompt
    /// </param>
    /// <returns>
    /// The picked point
    /// </returns>
    public static Point PickAPoint(string prompt = null)
    {
      Point pick = null;
      try
      {
          pick = (prompt == null) ? ModelPicker.PickPoint() : ModelPicker.PickPoint(prompt);
      }
      catch (Exception ex)
      {
          if (ex.Message.Contains("interrupt")) { }
          else Trace.WriteLine(ex.Message);
      }
      return pick;
    }
		
		public static Polygon PickTwoPoints(string prompt = null)
    {
      ArrayList pick = null;
      Polygon polygon = null;
      try
      {
        while (pick == null)
        {
            if (prompt == null) pick = ModelPicker.PickPoints(Picker.PickPointEnum.PICK_TWO_POINTS);
            else pick = ModelPicker.PickPoints(Picker.PickPointEnum.PICK_TWO_POINTS, prompt);
        }
        polygon = new Polygon();
        polygon.Points.AddRange(pick);
      }
      catch (Exception ex)
      {
        if (ex.Message.Contains("interrupt")) { }
        else Trace.WriteLine(ex.Message);
      }
      return polygon;
    }
  }
}
