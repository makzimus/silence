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
        private ToolStripMenuItem _connectItem;
        private ToolStripMenuItem _mutedProfilesItem;
        private ToolStripMenuItem _defaultProfilesItem;
        private readonly Timer _openRGBTimer = new() { Interval = 1000 };

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

            var openRGBItem = _contextMenu.Items.Add("OpenRGB") as ToolStripMenuItem;
            _connectItem = openRGBItem.DropDownItems.Add("Connect") as ToolStripMenuItem;
            _connectItem.ToolTipText = "Connect";
            _connectItem.Click += OnConnectPressed;

            var muteItem = (ToolStripMenuItem)_contextMenu.Items.Add("&Mute");
            muteItem.Checked = !_settings.PlayAudio;
            muteItem.Click += OnMutePressed;

            _defaultProfilesItem = openRGBItem.DropDownItems.Add("Default Profile") as ToolStripMenuItem;
            _mutedProfilesItem = openRGBItem.DropDownItems.Add("Muted Profile") as ToolStripMenuItem;

            var exitItem = _contextMenu.Items.Add("E&xit");
            exitItem.Click += OnExitPressed;

            _openRGBTimer.Tick += CheckOpenRGBConnection;
            TryConnectOpenRGB();

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
                    Console.WriteLine("OpenRGB connected.");
                    _connectItem.Checked = _openRGB.Connected;
                    _openRGBTimer.Start();
                    UpdateOpenRGBProfiles();
                }
                catch (TimeoutException)
                {
                    Console.WriteLine("OpenRGB failed to connect.");
                }
            }
        }

        private void CheckOpenRGBConnection(object sender, EventArgs e)
        {
            if (_openRGB != null)
            {
                try
                {
                    _openRGB.GetControllerCount();
                }
                catch (Exception)
                {
                    // Swallow any exceptions.
                }

                if (!_openRGB.Connected)
                {
                    OnOpenRGBDisconnected();
                }
            }
        }

        private void OnOpenRGBDisconnected()
        {
            Console.WriteLine("OpenRGB disconnected.");
            _openRGB?.Dispose();
            _openRGB = null;
            _connectItem.Checked = false;
            _openRGBTimer.Stop();
        }

        private void UpdateOpenRGBProfiles()
        {
            // Clear current items
            _mutedProfilesItem.DropDownItems.Clear();
            _defaultProfilesItem.DropDownItems.Clear();

            var profiles = _openRGB.GetProfiles();
            foreach (var profile in profiles)
            {
                var mutedItem = _mutedProfilesItem.DropDownItems.Add(profile);
                mutedItem.Name = profile;
                mutedItem.Click += OnProfileChanged;

                var defaultItem = _defaultProfilesItem.DropDownItems.Add(profile);
                defaultItem.Name = profile;
                defaultItem.Click += OnProfileChanged;
            }

            CheckDropDownMenuItem(_mutedProfilesItem, _settings.OpenRGBMutedProfile);
            CheckDropDownMenuItem(_defaultProfilesItem, _settings.OpenRGBDefaultProfile);
        }

        private static void CheckDropDownMenuItem(ToolStripMenuItem parent, string itemName)
        {
            if (parent.DropDownItems.ContainsKey(itemName))
            {
                var item = parent.DropDownItems[itemName] as ToolStripMenuItem;
                item.Checked = true;
            }
        }

        private void OnProfileChanged(object sender, EventArgs e)
        {
            var menuItem = sender as ToolStripMenuItem;
            var ownerItem = menuItem.OwnerItem as ToolStripMenuItem;

            if (ownerItem == _mutedProfilesItem)
            {
                _settings.OpenRGBMutedProfile = menuItem.Name;
                if (_muted)
                {
                    SelectProfile(menuItem.Name);
                }
            }
            else
            {
                _settings.OpenRGBDefaultProfile = menuItem.Name;
                if (!_muted)
                {
                    SelectProfile(menuItem.Name);
                }
            }

            _settings.Save();

            foreach (ToolStripMenuItem item in ownerItem.DropDownItems)
            {
                item.Checked = false;
            }

            menuItem.Checked = true;
        }

        private void ToggleRGB()
        {
            if (_openRGB == null)
            {
                return;
            }

            var profile = _muted ? _settings.OpenRGBMutedProfile : _settings.OpenRGBDefaultProfile;
            SelectProfile(profile);
        }

        private void SelectProfile(string profile)
        {
            try
            {
                _openRGB.LoadProfile(profile);
            }
            catch (Exception)
            {
                _openRGB.Dispose();
                _openRGB = null;
            }
        }

        private void OnIconClicked(object sender, MouseEventArgs e)
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
            var menuItem = sender as ToolStripMenuItem;
            menuItem.Checked = _settings.PlayAudio;
            _settings.PlayAudio = !_settings.PlayAudio;
            _settings.Save();
        }

        private void OnConnectPressed(object sender, EventArgs e)
        {
            if (_openRGB?.Connected == true)
            {
                OnOpenRGBDisconnected();
            }
            else
            {
                TryConnectOpenRGB();
            }
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
