// Generated by Tekla.Technology.Akit.ScriptBuilder



namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            akit.Callback("acmd_check_database", "1", "main_frame");
            akit.Callback("acmd_check_database", "XS_LIB_CORRECT", "main_frame");
            akit.Callback("acmd_partnumbers_all", "full", "main_frame");
            akit.PushButton("xs_save_button", "xs_report");
            akit.PushButton("warning_ok", "full_numbering_ok");
        }
    }
}
