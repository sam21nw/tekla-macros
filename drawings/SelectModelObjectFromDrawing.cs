// Generated by Tekla.Technology.Akit.ScriptBuilder
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
// Tekla Structures namespaces
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Solid;
using Tekla.Structures.Model.UI;
using TSM = Tekla.Structures.Model;
using TSMUI = Tekla.Structures.Model.UI;
using TS3D = Tekla.Structures.Geometry3d;
using TSD = Tekla.Structures.Drawing;
// Additional namespace references

// Notes

namespace Tekla.Technology.Akit.UserScript
{
  public static class Script
  {
    public static void Run(Tekla.Technology.Akit.IScript akit)
    {
      Model model = new Model();
      TSD.DrawingHandler drawingHandler = new TSD.DrawingHandler();
      TSD.Drawing drawing = drawingHandler.GetActiveDrawing();
      TSD.DrawingObjectEnumerator selectedObj = null;
      TSMUI.ModelObjectSelector selector = new TSMUI.ModelObjectSelector();
			ArrayList objects = new ArrayList();

      try
      {
        selectedObj = drawingHandler.GetDrawingObjectSelector().GetSelected();
        while (selectedObj.MoveNext())
        {
          if (selectedObj.Current is TSD.ModelObject)
          {
            TSD.ModelObject drawingPart = selectedObj.Current as TSD.ModelObject;
            TSM.ModelObject part = new TSM.Model().SelectModelObject(drawingPart.ModelIdentifier) as TSM.ModelObject;
            objects.Add(part);
          }
        }
        selector.Select(objects);
        model.CommitChanges();
				
				akit.Callback("acmdZoomToSelected", "", "main_frame");
      }
      catch { }
    }
  }
}
