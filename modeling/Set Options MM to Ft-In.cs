// Generated by Tekla.Technology.Akit.ScriptBuilder



namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            akit.Callback("acmdDisplayOptionSettingsDialog", "", "main_frame");
            akit.ListSelect("dia_option_settings", "listCategory", "Units and decimals");
            akit.ValueChange("dia_option_settings", "Optionmenu_10", "1005");
            akit.ValueChange("dia_option_settings", "Optionmenu_10", "1000");
            akit.ValueChange("dia_option_settings", "Optionmenu_10", "1005");
            akit.TabChange("dia_option_settings", "Container_957", "Container_960");
            akit.ValueChange("dia_option_settings", "Optionmenu_29", "1405");
            akit.ValueChange("dia_option_settings", "Optionmenu_29", "1400");
            akit.ModalDialog(1);
            akit.PushButton("butApply", "dia_option_settings");
            akit.PushButton("butCancel", "dia_option_settings");
        }
    }
}