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

        [UserScopedSetting]
        [DefaultSettingValue("Muted")]
        public string OpenRGBMutedProfile
        {
            get
            {
                return (string)this[nameof(OpenRGBMutedProfile)];
            }
            set
            {
                this[nameof(OpenRGBMutedProfile)] = value;
            }
        }

        [UserScopedSetting]
        [DefaultSettingValue("Default")]
        public string OpenRGBDefaultProfile
        {
            get
            {
                return (string)this[nameof(OpenRGBDefaultProfile)];
            }
            set
            {
                this[nameof(OpenRGBDefaultProfile)] = value;
            }
        }
    }
}