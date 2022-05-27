namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            akit.Callback("acmd_display_dialog", "dia_view_dialog", "main_frame");
            akit.PushButton("v1_filter", "dia_view_dialog");
            akit.PushButton("NewButton", "diaViewObjectGroup");
            akit.PushButton("dia_pa_modify", "diaViewObjectGroup");
            akit.PushButton("dia_pa_ok", "diaViewObjectGroup");
            akit.PushButton("v1_show", "dia_view_dialog");
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
            akit.ValueChange("Modelling view setup", "sd_part_display_type", "2");
            akit.ValueChange("Modelling view setup", "sd_object_display_type", "2");
            akit.ValueChange("Modelling view setup", "sd_joint_object_display_type", "2");
            akit.ValueChange("Modelling view setup", "sd_bolt_display_type", "2");
            akit.TabChange("Modelling view setup", "Container_647", "Container_650");
            akit.ValueChange("Modelling view setup", "sd_part_centerline", "0");
            akit.ValueChange("Modelling view setup", "PartReferenceLine", "0");
            akit.ValueChange("Modelling view setup", "sd_profile_text", "0");
            akit.ValueChange("Modelling view setup", "sd_joint_text", "0");
            akit.TabChange("Modelling view setup", "Container_647", "Container_648");
            akit.ValueChange("Modelling view setup", "sd_screws", "1");
            akit.ValueChange("Modelling view setup", "sd_joint_screws", "1");
            akit.PushButton("sd_modify", "Modelling view setup");
            akit.PushButton("sd_apply", "Modelling view setup");
            akit.PushButton("sd_ok", "Modelling view setup");
            akit.PushButton("v1_ok", "dia_view_dialog");
        }
    }
}
