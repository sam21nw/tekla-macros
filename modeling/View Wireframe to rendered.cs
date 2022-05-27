namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            akit.Callback("acmd_display_dialog", "dia_view_dialog", "main_frame");
            akit.ValueChange("dia_view_dialog", "v1_view_type", "1");
            akit.PushButton("v1_modify", "dia_view_dialog");
            akit.PushButton("v1_ok", "dia_view_dialog");
        }
		
    }
}
