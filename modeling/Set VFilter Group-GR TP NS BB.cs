#pragma warning disable 1633 // Unrecognized #pragma directive
#pragma reference "Tekla.Macros.Wpf.Runtime"
#pragma reference "Tekla.Macros.Akit"
#pragma reference "Tekla.Macros.Runtime"
#pragma warning restore 1633 // Unrecognized #pragma directive

namespace UserMacros {
    public sealed class Macro {
        [Tekla.Macros.Runtime.MacroEntryPointAttribute()]
        public static void Run(Tekla.Macros.Runtime.IMacroRuntime runtime) {
            Tekla.Macros.Akit.IAkitScriptHost akit = runtime.Get<Tekla.Macros.Akit.IAkitScriptHost>();
            Tekla.Macros.Wpf.Runtime.IWpfMacroHost wpf = runtime.Get<Tekla.Macros.Wpf.Runtime.IWpfMacroHost>();
            wpf.InvokeCommand("CommandRepository", "View.Properties");
            akit.PushButton("v1_filter", "dia_view_dialog");
            akit.ValueChange("diaViewObjectGroup", "get_menu", "Group - GR TP BB NS");
            akit.PushButton("dia_pa_modify", "diaViewObjectGroup");
            akit.PushButton("dia_pa_ok", "diaViewObjectGroup");
            akit.PushButton("v1_ok", "dia_view_dialog");
        }
    }
}
