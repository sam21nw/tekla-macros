using System;
using System.Diagnostics;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using Tekla.Structures.Model.UI;
using Tekla.BIM.Quantities;
using Tekla.Structures;
using Tekla.Structures.Model.Operations;
using System.Xml;
[assembly: Tekla.Technology.Scripting.Compiler.Reference("Tekla.BIM.Toolkit")]
[assembly: Tekla.Technology.Scripting.Compiler.Reference("System.Xml")]
namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(IScript akit)
        {
            try
            {
                InquirePointLocation.RunMacro(akit);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.InnerException + ex.Message + ex.StackTrace);
                Operation.DisplayPrompt(ex.InnerException + ex.Message);
            }
        }
    }

    public static class InquirePointLocation
    {

        public const string PickPointPrompt = "Pick point in model to inquire.";
        public const bool PaintTextInModel = true;
    		private static readonly Color TextColor = new Color(0.92, 0.11, 0.16); // red color

        /// <summary>
        /// Call to run create dimension logic
        /// </summary>
        /// <param name="akit">Akit for scripting if needed</param>
        public static void RunMacro(IScript akit)
        {
            //Prompt user for point in for 
            var pickedPoint = PickerService.PickAPoint(PickPointPrompt);
            if (pickedPoint == null) return;

            var translatedString = "Point Coordinates: ";
            translatedString += pickedPoint.GetFormattedString();

            //Display to user
            var insertPt = new Point(pickedPoint);
            insertPt.Translate(new Vector(1, 1, 0).GetNormal() * 5);
            if (PaintTextInModel) new Model().DrawText(translatedString, insertPt, TextColor);
            Operation.DisplayPrompt(translatedString);
        }
    }

    public static class MacroExtensions
    {
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
        /// Gets point to usimp units where needed and formats
        /// </summary>
        /// <param name="pt"></param>
        /// <returns></returns>
        public static string GetFormattedString(this Point pt)
        {
            var dx = new Length(pt.X);
            var dy = new Length(pt.Y);
            var dz = new Length(pt.Z);
            return string.Format("({0}, {1}, {2})", dx.ToCurrentUnits(), dy.ToCurrentUnits(), dz.ToCurrentUnits());
        }
    }

    /// <summary>
    /// Extensions class for Tekla.BIM.Quantities class
    /// </summary>
    public static class TsLength
    {
        /// <summary>
        /// Returns ft-fraction inch string rounded to 1/16 if XS_IMPERIAL=TRUE, mm to to decimal places otherwise
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
    }

    public static class PickerService
    {
        /// <summary>Tekla Structures Model Internal Picker </summary>
        public static readonly Picker ModelPicker = new Picker();

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
    }
}
