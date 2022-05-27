using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using Tekla.Structures;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;

[assembly: Tekla.Technology.Scripting.Compiler.Reference("Tekla.Application.Library")]
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
                OpenShopDrawing.RunMacro(akit);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.InnerException + ex.Message + ex.StackTrace);
            }
        }
    }

    public static class OpenShopDrawing
    {
        private static List<AssemblyDrawing> AssemblyDrawings { get; set; }
        private static List<CastUnitDrawing> CastUnitDrawings { get; set; }
        private static List<SinglePartDrawing> SinglePartDrawings { get; set; }
        private static DrawingHandler Handler { get { return new DrawingHandler(); } }

        public static void RunMacro(IScript akit)
        {
            var drawingOpened = false;

            //Get object from user
            var picker = new Tekla.Structures.Model.UI.Picker();
            var pickedPart = picker.PickObject(Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART) as Tekla.Structures.Model.Part;

            //Check if numbering is up to date
            if (!Operation.IsNumberingUpToDate(pickedPart))
                Operation.DisplayPrompt("Numbering is not up to date, run numbering first");

            //Check object from user
            if (pickedPart == null)
            {
                Operation.DisplayPrompt("Unable to get selected object.");
                return;
            }

            //Cache drawings from model
            if (!CacheDrawings()) return;

            //Get assembly mark
            var assemblyMark = string.Empty;
            pickedPart.GetReportProperty("ASSEMBLY_POS", ref assemblyMark);
            if (string.IsNullOrEmpty(assemblyMark))
            {
                Operation.DisplayPrompt(string.Format("Failed to get ASSEMBLY_POS for part {0}", pickedPart.Identifier.GUID));
                return;
            }

            //Get fake material type
            var fakeMaterialType = string.Empty;
            pickedPart.GetReportProperty("MATERIAL_TYPE", ref fakeMaterialType);
            if (string.IsNullOrEmpty(fakeMaterialType))
            {
                Operation.DisplayPrompt(string.Format("Failed to get MATERIAL_TYPE for part {0}", pickedPart.Identifier.GUID));
                return;
            }

            //Check first if cast unit type drawing
            if (fakeMaterialType == "CONCRETE")
            {
                foreach (var dg in CastUnitDrawings)
                {
                    var drawingMark = GetDrawingUsableMark(dg);
                    if (drawingMark != assemblyMark) continue;
                    if (!ForceUpdateIfNeeded(dg, akit)) return;
                    drawingOpened = Handler.SetActiveDrawing(dg, true);
                }
                if (!drawingOpened)
                    Operation.DisplayPrompt(string.Format("No drawing exists yet for selected part {0}", assemblyMark));
            }
            else //Drawing type is not concrete type, either wood or steel or miscellaneous
            {
                //First find if there is a single part drawing for selected
                SinglePartDrawing foundSinglePartDrawing = null;
                foreach (var dg in SinglePartDrawings)
                {
                    var drawingMark = GetDrawingUsableMark(dg);
                    if (drawingMark != assemblyMark) continue;
                    foundSinglePartDrawing = dg;
                    break;
                }

                //Second find if there is an Assembly part drawing for selected
                AssemblyDrawing foundAssemblyDrawing = null;
                foreach (var dg in AssemblyDrawings)
                {
                    var drawingMark = GetDrawingUsableMark(dg);
                    if (drawingMark != assemblyMark) continue;
                    foundAssemblyDrawing = dg;
                }

                if (foundSinglePartDrawing == null & foundAssemblyDrawing == null)
                {
                    Operation.DisplayPrompt(string.Format("No drawings exist yet for selected part {0}", assemblyMark));
                    return;
                }
                if (foundSinglePartDrawing == null) //Assembly drawing exits, but not single part
                {
                    if (!ForceUpdateIfNeeded(foundAssemblyDrawing, akit)) return;
                    drawingOpened = Handler.SetActiveDrawing(foundAssemblyDrawing);
                }
                else if (foundAssemblyDrawing == null) //Single part drawing exits, but not assembly part
                {
                    if (!ForceUpdateIfNeeded(foundSinglePartDrawing, akit)) return;
                    drawingOpened = Handler.SetActiveDrawing(foundSinglePartDrawing);
                }
                else //Both drawings exist, ask user which one to open
                {
                    var result = MessageBox.Show(TeklaStructures.MainWindow,
                        "Assembly and Single Part Drawings found:\nClick Yes to open the Assembly Drawing\nClick No to open the Single Part Drawing\nClick Cancel to abort", 
                        "Open Drawing Choice", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);

                    switch (result)
                    {
                        case DialogResult.Yes:
                            if (!ForceUpdateIfNeeded(foundAssemblyDrawing, akit)) return;
                            drawingOpened = Handler.SetActiveDrawing(foundAssemblyDrawing);
                            break;
                        case DialogResult.No:
                            if (!ForceUpdateIfNeeded(foundSinglePartDrawing, akit)) return;
                            drawingOpened = Handler.SetActiveDrawing(foundSinglePartDrawing);
                            break;
                        default:
                            Operation.DisplayPrompt(string.Format("Drawing open macro aborted by user"));
                            return;
                    }
                }

                //Check if any steel type drawing was found and opened
                if (!drawingOpened)
                    Operation.DisplayPrompt(string.Format("No drawing exists yet for selected part {0}", assemblyMark));
            }
        }

        private static string GetDrawingUsableMark(Drawing dg)
        {
            var realMark = string.Empty;
            if (dg is AssemblyDrawing)
            {
                var typeDrawing = (AssemblyDrawing)dg;
                var modelPart = new Model().SelectModelObject(typeDrawing.AssemblyIdentifier);
                modelPart.GetReportProperty("ASSEMBLY_POS", ref realMark);
                return realMark;
            }
            if (dg is CastUnitDrawing)
            {
                var typeDrawing = (CastUnitDrawing)dg;
                var modelPart = new Model().SelectModelObject(typeDrawing.CastUnitIdentifier);
                modelPart.GetReportProperty("ASSEMBLY_POS", ref realMark);
                return realMark;
            }
            if (dg is SinglePartDrawing)
            {
                var typeDrawing = (SinglePartDrawing)dg;
                var modelPart = new Model().SelectModelObject(typeDrawing.PartIdentifier);
                modelPart.GetReportProperty("ASSEMBLY_POS", ref realMark);
                return realMark;
            }
            return string.Empty;
        }

        private static bool ForceUpdateIfNeeded(Drawing dg, Tekla.Technology.Akit.IScript akit)
        {
            if (dg.UpToDateStatus == DrawingUpToDateStatus.DrawingIsUpToDate) return true;
            Trace.WriteLine("Drawing Not Up to Date: " + dg.UpToDateStatus);

            if (akit == null) return false;

            //Get identifier from drawing type
            Tekla.Structures.Identifier ident;
            if (dg is AssemblyDrawing) ident = ((AssemblyDrawing)dg).AssemblyIdentifier;
            else if (dg is CastUnitDrawing) ident = ((CastUnitDrawing) dg).CastUnitIdentifier;
            else if (dg is SinglePartDrawing) ident = ((SinglePartDrawing) dg).PartIdentifier;
            else throw new ArgumentOutOfRangeException();

            //Open drawing list and select object in model
            var selector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            var thisAssembly = new Model().SelectModelObject(ident);
            selector.Select(new ArrayList { thisAssembly });
            Tekla.Structures.ModelInternal.Operation.dotStartAction("DrawingList", null);

            //Macro Operations : Filter drawing list by selected, select 1st drawing, press update
            {
                akit.PushButton("dia_draw_filter_by_parts", "Drawing_selection");
                akit.TableSelect("Drawing_selection", "dia_draw_select_list", 1);
                akit.PushButton("dia_draw_select_update", "Drawing_selection");
            }

            //Drawing was able to be updated and opened
            if (dg.UpToDateStatus == DrawingUpToDateStatus.DrawingIsUpToDate)
            {
                Trace.WriteLine(string.Format("Drawing {0} updated due to out of date status", dg.Mark));
                return true;
            }

            //Drawing update failed, write error messages
            var msg = string.Format("Unable to update drawing {0}, macro cannot open out of date drawing", dg.Mark);
            Trace.WriteLine(msg);
            Operation.DisplayPrompt(msg);
            return false;
        }

        private static bool CacheDrawings()
        {
            //Clear old data
            CastUnitDrawings = new List<CastUnitDrawing>();
            AssemblyDrawings = new List<AssemblyDrawing>();
            SinglePartDrawings = new List<SinglePartDrawing>();

            //Get all drawings
            var drawingCollection = Handler.GetDrawings();
            if (drawingCollection.GetSize() < 1)
            {
                Operation.DisplayPrompt("No drawings are created yet.");
                return false;
            }

            //cache drawings locally
            foreach (var dg in drawingCollection)
            {
                if (dg is CastUnitDrawing && !CastUnitDrawings.Contains(dg as CastUnitDrawing))
                    CastUnitDrawings.Add(dg as CastUnitDrawing);
                else if (dg is AssemblyDrawing && !AssemblyDrawings.Contains(dg as AssemblyDrawing))
                    AssemblyDrawings.Add(dg as AssemblyDrawing);
                else if (dg is SinglePartDrawing && !SinglePartDrawings.Contains(dg as SinglePartDrawing))
                    SinglePartDrawings.Add(dg as SinglePartDrawing);
            }
            return true;
        }
    }

    /// <summary>
    /// Internal String helper class
    /// </summary>
    static class StringTools
    {
        /// <summary>
        /// Gets left side string from text
        /// </summary>
        /// <param name="text">Original text to search within</param>
        /// <param name="sub">Text to index position</param>
        /// <returns>Left handed side of string</returns>
        public static string Left(string text, string sub)
        {
            string result = string.Empty;
            int index = text.IndexOf(sub, System.StringComparison.Ordinal);
            if (index != -1)
                result = text.Substring(0, index).Trim();
            return result;
        }

        /// <summary>
        /// Gets right side string from text
        /// </summary>
        /// <param name="text">Original text to search within</param>
        /// <param name="sub">Text to index position</param>
        /// <returns>Right handed side of string</returns>
        public static string Right(string text, string sub)
        {
            string result = string.Empty;
            int index = text.IndexOf(sub, System.StringComparison.Ordinal);
            if (index != -1)
                result = text.Substring(index + 1, text.Length - index - 1).Trim();
            return result;
        }
    }
}
