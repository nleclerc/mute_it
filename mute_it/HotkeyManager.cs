using System;
using System.Windows.Forms;

namespace mute_it
{
    public static class Constants
    {
        public const int NOMOD = 0x0000;
        public const int ALT = 0x0001;
        public const int CTRL = 0x0002;
        public const int SHIFT = 0x0004;
        public const int WIN = 0x0008;

        public const int WM_HOTKEY_MSG_ID = 0x0312;
    }

    class HotkeyManager : NativeWindow, IDisposable
    {
        private MuteItContext context;  

        public HotkeyManager(MuteItContext context)
        {
            this.context = context;
            CreateHandle(new CreateParams());
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Constants.WM_HOTKEY_MSG_ID)
            {
                if (m.WParam.ToInt32() == MuteItContext.MUTE_CODE)
                {
                    context.muteMic();
                }

                if (m.WParam.ToInt32() == MuteItContext.UNMUTE_CODE)
                {
                    context.unmuteMic();
                }
            }
            base.WndProc(ref m);
        }

        public void Dispose()
        {
            DestroyHandle();
        }
    }
}
