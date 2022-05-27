namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            akit.Callback("acmd_display_dialog", "dia_view_dialog", "main_frame");
            akit.PushButton("v1_filter", "dia_view_dialog");
            akit.ValueChange("diaViewObjectGroup", "get_menu", "Group - GRAssPartsAll");
            akit.PushButton("dia_pa_modify", "diaViewObjectGroup");
            akit.PushButton("dia_pa_ok", "diaViewObjectGroup");
            akit.PushButton("v1_ok", "dia_view_dialog");
        }
    }
}
