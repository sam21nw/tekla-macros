namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            akit.Callback("acmd_save_model_attributes", "", "main_frame");
            akit.Callback("acmd_load_model_attributes", "", "main_frame");
        }
    }
}
