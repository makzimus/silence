using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using System.Media;
using CoreAudio;
using OpenRGB.NET;

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
        private readonly KeyboardHook _keyboardHook = new();
        private readonly Settings _settings = new();
        private OpenRGBClient _openRGB;

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

            TryConnectOpenRGB();

            var muteItem = (ToolStripMenuItem)_contextMenu.Items.Add("&Mute");
            muteItem.Checked = !_settings.PlayAudio;
            muteItem.Click += OnMutePressed;

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
            var hotkey = HotkeyParser.Parse(_settings.Hotkey);
            _keyboardHook.RegisterHotKey(hotkey.Item1, hotkey.Item2);

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
            if (playAudio && _settings.PlayAudio)
            {
                _audio[mute ? 0 : 1].Play();
            }

            ToggleRGB();
        }

        private void TryConnectOpenRGB()
        {
            if (_openRGB == null)
            {
                try
                {
                    _openRGB = new OpenRGBClient(name: "Silence", autoconnect: true, timeout: 1000);
                }
                catch (TimeoutException)
                {
                    // failed to connect
                }
            }
        }

        private void ToggleRGB()
        {
            if (_openRGB == null)
            {
                return;
            }

            var profile = _muted ? _settings.OpenRGBMutedProfile : _settings.OpenRGBDefaultProfile;
            _openRGB.LoadProfile(profile);
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

        private void OnExitPressed(object sender, EventArgs e)
        {
            _icon.Visible = false;
            Application.Exit();
        }

        private void OnMutePressed(object sender, EventArgs e)
        {
            var menuItem = (ToolStripMenuItem)sender;
            menuItem.Checked = _settings.PlayAudio;
            _settings.PlayAudio = !_settings.PlayAudio;
            _settings.Save();
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
