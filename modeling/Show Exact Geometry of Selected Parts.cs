// Generated by Tekla.Technology.Akit.ScriptBuilder
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Windows;


namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {

        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        const int KEY_DOWN_EVENT = 0x0000; //Key down flag
        const int KEY_UP_EVENT = 0x0002; //Key up flag

        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            // Press the Shift key
            keybd_event((byte)Keys.ShiftKey, 0, KEY_DOWN_EVENT, 0);
            try
            {
                //Perform action
                akit.Callback("acmd_draw_selected_hidden_lines", "own", "View_01 window_1");
            }
            finally
            {
                // Release the Shift key
                keybd_event((byte)Keys.ShiftKey, 0, KEY_UP_EVENT, 0);
            }

            SendKeys.SendWait("+");

        }
    }
}