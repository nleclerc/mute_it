using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using NAudio.CoreAudioApi;

namespace mute_it
{
    class MuteItContext : ApplicationContext
    {
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public static int MUTE_CODE = 123;
        public static int UNMUTE_CODE = 234;

        private static Double REFRESH_INTERVAL = 1000;
        private NotifyIcon tbIcon;
        private System.Timers.Timer refreshTimer;
        private HotkeyManager hkManager;

        public MuteItContext()
        {
            tbIcon = createIcon();
            updateMicStatus();
            refreshTimer = startTimer();
            hkManager = registerHotkeys();
        }

        private HotkeyManager registerHotkeys()
        {
            var manager = new HotkeyManager(this);

            RegisterHotKey(manager.Handle, MUTE_CODE, Constants.ALT + Constants.SHIFT, (int)Keys.P);
            RegisterHotKey(manager.Handle, UNMUTE_CODE, Constants.ALT + Constants.SHIFT, (int)Keys.O);

            return manager;
        }

        private NotifyIcon createIcon()
        {
            var exitMenuItem = new MenuItem("Exit", new EventHandler(handleExit));

            var icon = new NotifyIcon();
            icon.Icon = Properties.Resources.mic_off;
            icon.ContextMenu = new ContextMenu(new MenuItem[] {exitMenuItem});
            icon.DoubleClick += (source, e) => toggleMicStatus();

            icon.Visible = true;

            return icon;
        }

        private System.Timers.Timer startTimer()
        {
            var timer = new System.Timers.Timer(REFRESH_INTERVAL);
            timer.Elapsed += (source, e) => updateMicStatus();
            timer.Start();
            return timer;
        }

        public void setMicMuteStatus(bool doMute)
        {
            var device = getPrimaryMicDevice();

            if (device != null)
            {
                device.AudioEndpointVolume.Mute = doMute;
                updateMicStatus(device);
            }
            else
            {
                updateMicStatus(null);
            }
        }

        public void muteMic()
        {
            setMicMuteStatus(true);
        }

        public void unmuteMic()
        {
            setMicMuteStatus(false);
        }

        private void toggleMicStatus()
        {
            var device = getPrimaryMicDevice();

            if (device != null)
            {
                device.AudioEndpointVolume.Mute = !device.AudioEndpointVolume.Mute;
                updateMicStatus(device);
            }
            else
            {
                updateMicStatus(null);
            }
        }

        private void updateMicStatus()
        {
            var device = getPrimaryMicDevice();
            updateMicStatus(device);
            //System.GC.Collect();
        }

        private void updateMicStatus(MMDevice device)
        {
            if (device == null || device.AudioEndpointVolume.Mute == true)
                tbIcon.Icon = Properties.Resources.mic_off;
            else
                tbIcon.Icon = Properties.Resources.mic_on;

            disposeDevice(device);
        }

        private MMDevice getPrimaryMicDevice()
        {
            var enumerator = new MMDeviceEnumerator();
            var result = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications);
            enumerator.Dispose();

            tbIcon.Text = result.DeviceFriendlyName;

            return result;
        }

        private void disposeDevice(MMDevice device)
        {
            if (device != null)
            {
                device.AudioEndpointVolume.Dispose();
                device.Dispose();
            }
        }

        private void handleExit(object sender, EventArgs e)
        {
            tbIcon.Visible = false;
            refreshTimer.Stop();
            Application.Exit();
        }
    }
}
