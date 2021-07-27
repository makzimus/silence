using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using System.Media;
using CoreAudio;

namespace Silence
{
    class Program
    {
        private NotifyIcon _icon;
        private ContextMenuStrip _contextMenu;
        private IContainer _components;
        private bool _muted;
        private readonly List<System.Drawing.Icon> _icons = new();
        private readonly List<SoundPlayer> _audio = new();
        private KeyboardHook _keyboardHook = new();

        private static Stream GetEmbeddedResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            return assembly.GetManifestResourceStream(resourceName);
        }

        private void LoadResources()
        {
            using (var stream = GetEmbeddedResource("Silence.res.mic_off.ico"))
            {
                _icons.Add(new System.Drawing.Icon(stream));
            }
            using (var stream = GetEmbeddedResource("Silence.res.mic_on.ico"))
            {
                _icons.Add(new System.Drawing.Icon(stream));
            }

            void loadClip(string clip)
            {
                using var stream = GetEmbeddedResource(clip);
                var snd = new SoundPlayer(stream);
                snd.Load();
                _audio.Add(snd);
            };
            loadClip("Silence.res.bloop.wav");
            loadClip("Silence.res.beep.wav");
        }

        private void Initialize()
        {
            _components = new Container();
            _contextMenu = new ContextMenuStrip(_components);
            var exitItem = _contextMenu.Items.Add("E&xit");
            exitItem.Click += OnExitPressed;

            LoadResources();
            _icon = new NotifyIcon(_components)
            {
                Icon = _icons[1],
                Text = "Silence",
                Visible = true
            };
            _icon.MouseClick += OnIconClicked;
            _icon.ContextMenuStrip = _contextMenu;

            _keyboardHook.KeyPressed += OnHotkeyPressed;
            _keyboardHook.RegisterHotKey(ModifierKeys.Control | ModifierKeys.Alt, Keys.N);

            MuteDevices(_muted, false);
        }

        private void MuteDevices(bool mute, bool playAudio = true)
        {
            _muted = mute;
            var deviceEnumerator = new MMDeviceEnumerator();
            var devices = deviceEnumerator.EnumerateAudioEndPoints(EDataFlow.eCapture, DEVICE_STATE.DEVICE_STATE_ACTIVE);
            foreach (var device in devices)
            {
                device.AudioEndpointVolume.Mute = mute;
            }

            _icon.Icon = _icons[mute ? 0 : 1];
            if (playAudio)
            {
                _audio[mute ? 0 : 1].Play();
            }
        }

        private void OnIconClicked(object Sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                MuteDevices(!_muted);
            }
        }

        void OnHotkeyPressed(object sender, KeyPressedEventArgs e)
        {
            MuteDevices(!_muted);
        }

        private void OnExitPressed(object Sender, EventArgs e)
        {
            _icon.Visible = false;
            Application.Exit();
        }

        [STAThread]
        static void Main()
        {
            var program = new Program();
            program.Initialize();
            Application.Run();
        }
    }
}
