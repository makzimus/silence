using System;
using System.Windows.Forms;

namespace Silence
{
    public static class HotkeyParser
    {
        private static (ModifierKeys, Keys) Default = (ModifierKeys.Ctrl | ModifierKeys.Alt, Keys.N);

        public static (ModifierKeys, Keys) Parse(string keyStr)
        {
            if (!string.IsNullOrEmpty(keyStr))
            {
                // Split.
                var split = keyStr.Split('+');

                // Parse modifiers.
                ModifierKeys modifiers = 0;
                for (int i = 0; i < split.Length - 1; ++i)
                {
                    if (Enum.TryParse(typeof(ModifierKeys), split[i], true, out var result))
                    {
                        modifiers |= (ModifierKeys)result;
                    }
                }

                // Parse key.
                if (Enum.TryParse(typeof(Keys), split[^1], true, out var keyObject))
                {
                    return (modifiers, (Keys)keyObject);
                }

                Console.WriteLine("Failed to parse hotkey.");
            }

            return Default;
        }
    }
}