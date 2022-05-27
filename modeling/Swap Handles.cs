//
// ###########################################################################################
// ### Name          : Swap_Handles Macro
// ### Version       : 2.0 for V19.0
// ###               : Visual C# / Visual Studio 2013
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
using System.Collections;
using System.Diagnostics;
using Tekla.Structures.Model;

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
                new SwapHandles();
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
                new SwapHandles();
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message + ex.StackTrace);
            }
        }
    }

    public class SwapHandles
    {
        public SwapHandles()
        {
            //Get selected objects and put them in an enumerator/container
            var selector = new Tekla.Structures.Model.UI.ModelObjectSelector();
            var myEnum = selector.GetSelectedObjects();

            //Cycle through selected objects
            while (myEnum.MoveNext())
            {
                //Cast beam
                if (myEnum.Current is Tekla.Structures.Model.Beam)
                {
                    var myBeam = myEnum.Current as Beam;

                    // Get part current handles
                    var startPoint = myBeam.StartPoint;
                    var endPoint = myBeam.EndPoint;

                    // Switch part handles
                    myBeam.StartPoint = endPoint;
                    myBeam.EndPoint = startPoint;

                    //Swap uda's for design forces
                    SwapEndForces(myBeam);

                    // modify beam and refresh model + undo 
                    myBeam.Modify();
                }
                else if(myEnum.Current is Tekla.Structures.Model.PolyBeam)
                {
                    var myBeam = myEnum.Current as PolyBeam;

                    // Get part current handles
                    var newPoints = new ArrayList();
                    var oldPoints = myBeam.Contour.ContourPoints;

                    //Copy points to new seperate list first
                    foreach (var cp in oldPoints)
                        newPoints.Add(cp);
                    newPoints.Reverse();

                    //Swap uda's for design forces
                    SwapEndForces(myBeam);

                    // modify beam and refresh model + undo 
                    myBeam.Contour.ContourPoints = newPoints;
                    myBeam.Modify();
                }
            }

            //Update model with changes
            new Model().CommitChanges();
        }

        private static void SwapEndForces(ModelObject myBeam)
        {
            var originalEnd1 = string.Empty;
            var originalEnd2 = string.Empty;
            myBeam.GetUserProperty("BM_FORCE1", ref originalEnd1);
            myBeam.GetUserProperty("BM_FORCE2", ref originalEnd2);
            myBeam.SetUserProperty("BM_FORCE1", originalEnd2);
            myBeam.SetUserProperty("BM_FORCE2", originalEnd1);
        }
    }
}
