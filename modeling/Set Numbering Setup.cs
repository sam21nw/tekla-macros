namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            akit.Callback("acmd_display_partnumbers_set_options", "", "main_frame");
            akit.ValueChange("m3_dialog", "m3_renumber_all", "0");
            akit.ValueChange("m3_dialog", "m3_use_std_parts", "0");
            akit.ValueChange("m3_dialog", "m3_use_old_numbers", "1");
            akit.ValueChange("m3_dialog", "m3_holes", "1");
            akit.ValueChange("m3_dialog", "m3_profile_name", "1");
            akit.ValueChange("m3_dialog", "m3_beam_orientation", "1");
            akit.ValueChange("m3_dialog", "m3_column_orientation", "0");
            akit.ValueChange("m3_dialog", "checkbutton_3664", "0");
            akit.ValueChange("m3_dialog", "checkbutton_3664_1", "0");
            akit.ValueChange("m3_dialog", "checkbutton_3664_2", "0");
            akit.ValueChange("m3_dialog", "checkbutton_3664_3", "0");
            akit.ValueChange("m3_dialog", "m3_tolerance_1", "0.800000000000");
            akit.ValueChange("m3_dialog", "m3_tolerance_2", "0.800000000000");
            akit.ValueChange("m3_dialog", "m3_tolerance_3", "0.800000000000");
            akit.ValueChange("m3_dialog", "m3_tolerance", "0.800000000000");
            akit.ValueChange("m3_dialog", "m3_new_options_om", "0");
            akit.ValueChange("m3_dialog", "m3_modified_options_om", "2");
            akit.ValueChange("m3_dialog", "m3_automatic_cloning", "1");
            akit.ValueChange("m3_dialog", "m3_save_numbering_save", "1");
            akit.ValueChange("m3_dialog", "m3_save_numbering_save", "0");
            akit.ValueChange("m3_dialog", "m3_SortBy1", "4");
            akit.ValueChange("m3_dialog", "m3_SortByUDA1", "USER_PHASE");
            akit.ValueChange("m3_dialog", "radiobox_11900", "1");
            akit.ValueChange("m3_dialog", "m3_SortBy2", "0");
            akit.ValueChange("m3_dialog", "get_menu", "standard");
            akit.ValueChange("m3_dialog", "saveas_file", "");
            akit.PushButton("attrib_save", "m3_dialog");
            akit.PushButton("m3_apply", "m3_dialog");
        }
    }
}
