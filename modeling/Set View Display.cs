using System;
using System.Diagnostics;
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;
using Tekla.Structures.Model.UI;

[assembly: Tekla.Technology.Scripting.Compiler.ReferenceAttribute("Tekla.Application.Library")]

namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
			bool check = XSImperialInputCheck();		
		
			akit.Callback("acmd_display_dialog", "dia_view_dialog", "main_frame");
			akit.PushButton("v1_show", "dia_view_dialog");
            akit.ValueChange("dia_view_dialog", "v1_view_type", "1");
			akit.ValueChange("Modelling view setup", "sd_parts", "1");
            akit.ValueChange("Modelling view setup", "sd_joint_parts", "1");
            akit.ValueChange("Modelling view setup", "sd_screws", "1");
            akit.ValueChange("Modelling view setup", "sd_joint_screws", "1");
            akit.ValueChange("Modelling view setup", "sd_object_display_type", "1");
            akit.ValueChange("Modelling view setup", "sd_joint_object_display_type", "4");
            akit.ValueChange("Modelling view setup", "sd_holes", "1");
            akit.ValueChange("Modelling view setup", "sd_joint_holes", "1");
            akit.ValueChange("Modelling view setup", "sd_weldings", "0");
            akit.ValueChange("Modelling view setup", "sd_joint_weldings", "0");
            akit.ValueChange("Modelling view setup", "chkButUserPlaneVisibility", "0");
            akit.ValueChange("Modelling view setup", "chkButUserPlaneVisibInJoints", "0");
            akit.ValueChange("Modelling view setup", "chkButRebarVisibility", "0");
            akit.ValueChange("Modelling view setup", "chkButRebarVisibilityInJoints", "0");
            akit.ValueChange("Modelling view setup", "chkButSurfacingVisibility", "0");
            akit.ValueChange("Modelling view setup", "chkButPourBreakVisibility", "0");
            akit.ValueChange("Modelling view setup", "chkButLoadVisibility", "0");
            akit.ValueChange("Modelling view setup", "sd_cuts", "0");
            akit.ValueChange("Modelling view setup", "sd_joint_cuts", "0");
            akit.ValueChange("Modelling view setup", "sd_fittings", "0");
            akit.ValueChange("Modelling view setup", "sd_joint_fittings", "0");
            akit.ValueChange("Modelling view setup", "sd_joint_symbols", "0");
            akit.ValueChange("Modelling view setup", "sd_joint_joint_symbols", "0");
            akit.ValueChange("Modelling view setup", "sd_construction_lines", "0");
            akit.ValueChange("Modelling view setup", "sd_reference_models", "0");
            akit.ValueChange("Modelling view setup", "sd_points", "0");
            akit.ValueChange("Modelling view setup", "sd_object_display_type", "1");
            akit.ValueChange("Modelling view setup", "sd_part_display_type", "2");
            akit.ValueChange("Modelling view setup", "sd_bolt_display_type", "2");
            akit.ValueChange("Modelling view setup", "sd_hole_display_type", "3");
            akit.TabChange("Modelling view setup", "Container_647", "Container_650");
            akit.ValueChange("Modelling view setup", "sd_part_centerline", "0");
            akit.ValueChange("Modelling view setup", "PartReferenceLine", "0");
            akit.ValueChange("Modelling view setup", "sd_profile_text", "0");
            akit.ValueChange("Modelling view setup", "sd_joint_text", "0");
            akit.ValueChange("Modelling view setup", "sd_grids", "1");
            akit.TabChange("Modelling view setup", "Container_647", "Container_648");
            akit.PushButton("sd_modify", "Modelling view setup");
            akit.PushButton("sd_ok", "Modelling view setup");
			
			if(check) {
            	akit.ValueChange("dia_view_dialog", "v1_depth_up", "6");
            	akit.ValueChange("dia_view_dialog", "v1_depth_down", "6");
            	akit.PushButton("v1_modify", "dia_view_dialog");
            	akit.PushButton("v1_cancel", "dia_view_dialog");
			}
			else {
				akit.ValueChange("dia_view_dialog", "v1_depth_up", "150.000");
            	akit.ValueChange("dia_view_dialog", "v1_depth_down", "150.000");
            	akit.PushButton("v1_modify", "dia_view_dialog");
            	akit.PushButton("v1_cancel", "dia_view_dialog");
			}
    }
	
	public static bool XSImperialInputCheck()
    	{
		    bool output = false;
		    string xsImperial = string.Empty;
		    TeklaStructuresSettings.GetAdvancedOption("XS_IMPERIAL_INPUT", ref xsImperial);
		    if (xsImperial == "TRUE")
		    {
		      output = true;
		    }
		    return output;
    	}
	}
}
