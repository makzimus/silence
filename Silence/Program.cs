using CoreAudio;
using Microsoft.Win32;
using OpenRGB.NET;
using OpenRGB.NET.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Media;
using System.Reflection;
using System.Windows.Forms;

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

        public bool OpenRGBConnected => _openRGB != null && _openRGB.Connected;

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
            SystemEvents.SessionSwitch += OnSystemSessionSwitch;

            _components = new Container();
            _contextMenu = new ContextMenuStrip(_components);

            var openRGBItem = _contextMenu.Items.Add("OpenRGB") as ToolStripMenuItem;
            _connectItem = openRGBItem.DropDownItems.Add("Connect") as ToolStripMenuItem;
            _connectItem.ToolTipText = "Connect";
            _connectItem.Click += OnConnectPressed;

            var hotkeyItem = (ToolStripMenuItem)_contextMenu.Items.Add("Set Hotkey");
            hotkeyItem.Click += OnHotkeyItemPressed;

            var muteItem = (ToolStripMenuItem)_contextMenu.Items.Add("&Mute");
            muteItem.Checked = !_settings.PlayAudio;
            muteItem.Click += OnMutePressed;

            _defaultProfilesItem = openRGBItem.DropDownItems.Add("Default Profile") as ToolStripMenuItem;
            _mutedProfilesItem = openRGBItem.DropDownItems.Add("Muted Profile") as ToolStripMenuItem;

            var gitItem = _contextMenu.Items.Add("Open Github");
            gitItem.Click += (s, e) => {
                const string url = "https://github.com/makzimus/silence";
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            };

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
            RegisterHotkey();

            MuteDevices(_muted, false);
        }

        private void RegisterHotkey()
        {
            var hotkey = HotkeyParser.Parse(_settings.Hotkey);
            _keyboardHook.ClearRegisteredHotkeys();
            _keyboardHook.RegisterHotKey(hotkey.Item1, hotkey.Item2);
        }

        private void MuteDevices(bool mute, bool playAudio = true, bool updateRGB = true)
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

            if (updateRGB)
            {
                UpdateOpenRGBProfile();
            }
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

        /// <summary>
        /// Selects the OpenRGB profile according to the current mute state.
        /// </summary>
        private void UpdateOpenRGBProfile()
        {
            if (!OpenRGBConnected)
            {
                return;
            }

            var profile = _muted ? _settings.OpenRGBMutedProfile : _settings.OpenRGBDefaultProfile;
            SelectProfile(profile);
        }

        /// <summary>
        /// Select an OpenRGB profile.
        /// </summary>
        /// <param name="profile">The profile to select.</param>
        private void SelectProfile(string profile)
        {
            try
            {
                _openRGB.LoadProfile(profile);
            }
            catch (Exception)
            {
                OnOpenRGBDisconnected();
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

        private void OnHotkeyItemPressed(object sender, EventArgs e)
        {
            KeyCaptureForm keyCaptureForm = new(_settings.Hotkey);
            if (keyCaptureForm.ShowDialog() == DialogResult.OK)
            {
                string keyCombination = keyCaptureForm.KeyLabel.Text;
                _settings.Hotkey = keyCombination;
                _settings.Save();
                RegisterHotkey();
            }
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

        private void OnSystemSessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                MuteDevices(true, playAudio: false, updateRGB: false);

                if (OpenRGBConnected)
                {
                    int controllerCount = _openRGB.GetControllerCount();
                    for (int i = 0; i < controllerCount; ++i)
                    {
                        var controllerData = _openRGB.GetControllerData(i);
                        var offArray = Enumerable.Repeat(new Color(), controllerData.Colors.Length).ToArray();
                        _openRGB.UpdateLeds(i, offArray);
                    }
                }
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                UpdateOpenRGBProfile();
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
