//
// ###########################################################################################
// ### Name          : Reopen_Drawing Macro
// ### Version       : 1.0 for V18.0
// ###               : Visual C# / Visual Studio 2010
// ### Created       : February 2012
// ### Modified      : 
// ### Author        : Charles Pool
// ### Released      : 
// ### Comment       :
// ### Description   : Simple script to swap beam position handles of selected objects
// ###                 
// ###                 
// ###########################################################################################
//

using System;
using System.Diagnostics;
using Tekla.Structures.Drawing;

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
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            try
            {
                new ReopenDrawing();
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message + ex.StackTrace);
            }
        }

        /// <summary>
        /// Internal method for debugging in console application
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            try
            {
                new ReopenDrawing();
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message + ex.StackTrace);
            }
        }
    }

    public class ReopenDrawing
    {
        public ReopenDrawing()
        {
            var drawingHandler = new DrawingHandler();
            var activeDrawing = drawingHandler.GetActiveDrawing();

            if (activeDrawing != null)
            {
                drawingHandler.SaveActiveDrawing();
                drawingHandler.SetActiveDrawing(activeDrawing, true);
                return;
            }
            Tekla.Structures.Model.Operations.Operation.DisplayPrompt("No drawing found, macro failed.");
        }
    }
}
