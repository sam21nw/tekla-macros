namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            akit.ValueChange("main_frame", "sel_filter", "Group - GRAssPartsAll");
        }
    }
}