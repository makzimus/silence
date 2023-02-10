using System;
using System.Configuration;

namespace Silence
{
    public class Settings : ApplicationSettingsBase
    {
        [UserScopedSetting]
        [DefaultSettingValue("true")]
        public bool PlayAudio
        {
            get
            {
                return (bool)this[nameof(PlayAudio)];
            }
            set
            {
                this[nameof(PlayAudio)] = value;
            }
        }

        [UserScopedSetting]
        [DefaultSettingValue("Ctrl+Alt+N")]
        public string Hotkey
        {
            get
            {
                return (string)this[nameof(Hotkey)];
            }
            set
            {
                this[nameof(Hotkey)] = value;
            }
        }
    }
}