using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing;

namespace Silence
{
    public class Native
    {
        public const int WH_KEYBOARD_LL = 13;
        public const int WM_KEYDOWN = 0x0100;

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT
        {
            public Keys key;
            public int scanCode;
            public int flags;
            public int time;
            public IntPtr extraInfo;
        }

        public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
    }

    public class KeyCaptureForm : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, Native.LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        private Label _keyLabel;
        private Button _confirmButton;
        private IntPtr _hookHandle = IntPtr.Zero;
        private bool _hasModifiers = false;

        public Label KeyLabel { get => _keyLabel; set => _keyLabel = value; }

        public KeyCaptureForm(string currentValue)
        {
            InitializeComponent(currentValue);
        }

        private void InitializeComponent(string currentValue)
        {
            KeyLabel = new Label();
            _confirmButton = new Button();
            SuspendLayout();
            // 
            // promptLabel
            // 
            Label promptLabel = new()
            {
                Text = "Press Key Combination:",
                AutoSize = true,
                TabIndex = 0
            };
            // 
            // keyText
            // 
            _keyLabel.AutoSize = true;
            _keyLabel.Name = "KeyLabel";
            _keyLabel.Text = currentValue;
            _keyLabel.MinimumSize = new Size(0, 20);
            // 
            // confirmButton
            // 
            _confirmButton.Name = "ConfirmButton";
            _confirmButton.MinimumSize = new Size(75, 23);
            _confirmButton.MaximumSize = _confirmButton.MinimumSize;
            _confirmButton.Text = "OK";
            _confirmButton.UseVisualStyleBackColor = true;
            _confirmButton.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _confirmButton.TextAlign = ContentAlignment.MiddleCenter;
            _confirmButton.Click += OnButtonClicked;
            // 
            // layoutPanel
            // 
            FlowLayoutPanel layoutPanel = new()
            {
                Dock = DockStyle.Fill,
                WrapContents = false,
                FlowDirection = FlowDirection.TopDown
            };
            layoutPanel.Controls.Add(promptLabel);
            layoutPanel.Controls.Add(_keyLabel);
            layoutPanel.Controls.Add(_confirmButton);

            ClientSize = new Size(200, 70);
            Controls.Add(layoutPanel);
            FormClosing += OnFormClosing;
            Name = "KeyCaptureForm";
            Text = "Set Hotkey";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Width = 200;
            ResumeLayout(false);
            PerformLayout();

            _hookHandle = SetWindowsHookEx(Native.WH_KEYBOARD_LL, LowLevelKeyboardProc, IntPtr.Zero, 0);
        }

        private IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)Native.WM_KEYDOWN)
            {
                Native.KBDLLHOOKSTRUCT hookStruct = (Native.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Native.KBDLLHOOKSTRUCT));

                string modifiers = string.Empty;
                if ((ModifierKeys & Keys.Alt) == Keys.Alt)
                {
                    modifiers += "Alt+";
                }
                if ((ModifierKeys & Keys.Control) == Keys.Control)
                {
                    modifiers += "Ctrl+";
                }
                if ((ModifierKeys & Keys.Shift) == Keys.Shift)
                {
                    modifiers += "Shift+";
                }

                _hasModifiers = modifiers != string.Empty;
                _keyLabel.Text = modifiers + hookStruct.key.ToString();
            }

            return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            UnhookWindowsHookEx(_hookHandle);
        }

        private void OnButtonClicked(object sender, EventArgs e)
        {
            DialogResult = _hasModifiers ? DialogResult.OK : DialogResult.Cancel;
            Close();
        }
    }
}