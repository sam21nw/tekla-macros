namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            akit.Callback("acmd_display_dialog", "dia_view_dialog", "main_frame");
            akit.PushButton("v1_show", "dia_view_dialog");
            akit.ValueChange("Modelling view setup", "sd_points", "1");
            akit.ValueChange("Modelling view setup", "sd_construction_lines", "0");
            akit.PushButton("sd_modify", "Modelling view setup");
            akit.PushButton("sd_ok", "Modelling view setup");
            akit.PushButton("v1_ok", "dia_view_dialog");
        }
    }
}
