using System.IO;
using Silence.Localization;
using Silence.Macro;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace Silence
{
    public partial class FormMain : Form
    {
        private const string ConfigurationFilePath = "config.json";

        private readonly MacroRecorder _recorder = new MacroRecorder();
        private readonly MacroPlayer _player = new MacroPlayer();

        private readonly ConfigurationFile _config;
        private readonly LanguagePack _languages;

        public FormMain()
        {
            InitializeComponent();

            if (!File.Exists(ConfigurationFilePath))
                new ConfigurationFile().Save(ConfigurationFilePath);
            _config = ConfigurationFile.FromFile(ConfigurationFilePath);

            _languages = new LanguagePack(@"lang");
            _languages.SelectLanguage(_config.LanguageCode);

            this.Load += FormMain_Load;
            this.FormClosing += FormMain_FormClosing;
        }

        private void recordControlButton_Click(object sender, EventArgs e) => StartRecording();

        private void stopControlButton_Click(object sender, EventArgs e) => StopRecording();

        private void playControlButton_Click(object sender, EventArgs e) => PlayRecording();

        private void HotkeyStartRecording_Pressed(object sender, EventArgs e) => StartRecording();

        private void HotkeyStopRecording_Pressed(object sender, EventArgs e) => StopRecording();

        private void HotkeyPlayRecording_Pressed(object sender, EventArgs e) => PlayRecording();

        private void clearControlButton_Click(object sender, EventArgs e)
        {
            // Confirm action.
            if (_recorder.CurrentMacro != null && _recorder.CurrentMacro.Events.Length > 0)
            {
                var result = MessageBox.Show(_languages.GetLocalizedString("confirm_clear_message"),
                    _languages.GetLocalizedString("confirm_clear_title"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    _recorder.Clear();
                    label_loadedFile.Text = "No Open File";
                }
            }
        }

        private void openControlButton_Click(object sender, EventArgs e)
        {
            // Confirm action.
            if (_recorder.CurrentMacro != null && _recorder.CurrentMacro.Events.Length > 0)
            {
                var result = MessageBox.Show(_languages.GetLocalizedString("confirm_open_message"),
                    _languages.GetLocalizedString("confirm_clear_title"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                    return;
            }

            // Browse for file
            var dialog = new OpenFileDialog
            {
                Title = _languages.GetLocalizedString("dialog_open_macro_title"),
                Filter = _languages.GetLocalizedString("dialog_open_macro_filter")
            };

            // Load macro into recorder.
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var loadedMacro = new Macro.Macro();
                loadedMacro.Load(dialog.FileName);
                label_loadedFile.Text = dialog.FileName;
                _recorder.LoadMacro(loadedMacro);
            }
        }

        private void saveControlButton_Click(object sender, EventArgs e)
        {
            // Check there is a macro to save.
            if (_recorder.CurrentMacro == null || _recorder.CurrentMacro.Events.Length == 0)
            {
                MessageBox.Show(_languages.GetLocalizedString("error_save_nomacro_message"),
                    _languages.GetLocalizedString("error_save_nomacro_title"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Choose file to save to.
            var dialog = new SaveFileDialog
            {
                Title = _languages.GetLocalizedString("dialog_save_macro_title"),
                Filter = _languages.GetLocalizedString("dialog_save_macro_filter")
            };

            // Save file.
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _recorder.CurrentMacro.Save(dialog.FileName);
                label_loadedFile.Text = dialog.FileName;
            }
        }

        private void loopControlButton_Click(object sender, EventArgs e)
        {
            // Set number of repetitions on player.
            var dialog = new RepetitionsDialog { Repetitions = _player.Repetitions };
            if (dialog.ShowDialog() == DialogResult.OK)
                _player.Repetitions = dialog.Repetitions;
        }

        private void StartRecording()
        {
            if (_player.IsPlaying)
                _player.CancelPlayback();

            // Confirm action.
            if (_recorder.CurrentMacro != null && _recorder.CurrentMacro.Events.Length > 0)
            {
                var result = MessageBox.Show(_languages.GetLocalizedString("confirm_append_message"),
                    _languages.GetLocalizedString("confirm_append_title"), MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                    _recorder.Clear();
                else if (result == DialogResult.Cancel)
                    return;
            }

            // Begin recording.
            _recorder.RecordMouse = chkRecordMouse.Checked;
            _recorder.StartRecording();
        }

        private void StopRecording()
        {
            if (_player.IsPlaying)
                _player.CancelPlayback();

            if (_recorder.IsRunning)
            {
                _recorder.StopRecording();

                if (_recorder.CurrentMacro != null && _recorder.CurrentMacro.Events.Length > 0)
                {
                    // Remove the initial and ending hotkeys if they exist
                    RemoveHotkeyEventsFromCurrentMacro(new int[] { (int)Keys.ControlKey, (int)Keys.ShiftKey, (int)Keys.F5 }, true);
                    RemoveHotkeyEventsFromCurrentMacro(new int[] { (int)Keys.LControlKey, (int)Keys.LShiftKey, (int)Keys.F5 }, true);
                    RemoveHotkeyEventsFromCurrentMacro(new int[] { (int)Keys.ControlKey, (int)Keys.ShiftKey, (int)Keys.F6 }, false);
                    RemoveHotkeyEventsFromCurrentMacro(new int[] { (int)Keys.LControlKey, (int)Keys.LShiftKey, (int)Keys.F6 }, false);
                }
            }
        }

        private void RemoveHotkeyEventsFromCurrentMacro(int[] findKeys, bool fromStart)
        {
            HashSet<int> foundKeys = new HashSet<int>();

            for (int i = 0; i < _recorder.CurrentMacro.EditableEvents.Count; i++)
            {
                var foundNotKey = false;
                var index = fromStart ? i : _recorder.CurrentMacro.EditableEvents.Count - 1 - i;
                var currentEvent = _recorder.CurrentMacro.EditableEvents[index];

                if (currentEvent is MacroKeyEvent keyEvent)
                {
                    if (findKeys.Contains(keyEvent.VirtualKeyCode))
                        foundKeys.Add(keyEvent.VirtualKeyCode);
                    else
                        foundNotKey = true;
                }
                else if (currentEvent is MacroDelayEvent || currentEvent is MacroMouseMoveEvent)
                {
                    // ignore for now
                }
                else
                {
                    foundNotKey = true;
                }

                if (foundNotKey)
                {
                    bool allFound = true;

                    for (int j = 0; j < findKeys.Length; j++)
                        if (!foundKeys.Contains(findKeys[j]))
                            allFound = false;

                    if (allFound)
                    {
                        // remove up to this
                        for (int j = 0; j < i; j++)
                        {
                            var removeIndex = fromStart ? j : _recorder.CurrentMacro.EditableEvents.Count - 1 - j;
                            var removeEvent = _recorder.CurrentMacro.EditableEvents[removeIndex];

                            if (removeEvent is MacroKeyEvent removeKeyEvent)
                            {
                                if (findKeys.Contains(removeKeyEvent.VirtualKeyCode))
                                {
                                    _recorder.CurrentMacro.EditableEvents.RemoveAt(removeIndex);
                                    j--;
                                    i--;
                                }
                            }
                        }
                    }

                    break;
                }
            }
        }

        private void PlayRecording()
        {
            if (_player.IsPlaying)
                _player.CancelPlayback();

            if (_recorder.IsRunning)
            {
                MessageBox.Show("Cannot start playing while recording!");
                return;
            }

            // Load and play macro.
            if (_recorder.CurrentMacro == null || _recorder.CurrentMacro.Events.Length == 0)
            {
                MessageBox.Show("No events recorded!");
                return;
            }

            _player.LoadMacro(_recorder.CurrentMacro);
            _player.PlayMacroAsync();
        }

        private HotKeyRegister hotkeyStartRecording = null;
        private HotKeyRegister hotkeyStopRecording = null;
        private HotKeyRegister hotkeyPlayRecording = null;

        private void FormMain_Load(object sender, EventArgs e)
        {
            // Load theme color for buttons.
            foreach (var control in panel1.Controls)
            {
                if (control is ControlButton button)
                {
                    button.MouseOutBackgroundColor = _config.ThemeColor.ToColor();
                    button.MouseOverBackgroundColor = _config.ThemeColor.ToColor(32);
                    button.MouseDownBackgroundColor = _config.ThemeColor.ToColor(-32);
                }
            }

            hotkeyStartRecording = new HotKeyRegister(this.Handle, 51, KeyModifiers.Control | KeyModifiers.Shift, Keys.F5);
            hotkeyStopRecording = new HotKeyRegister(this.Handle, 52, KeyModifiers.Control | KeyModifiers.Shift, Keys.F6);
            hotkeyPlayRecording = new HotKeyRegister(this.Handle, 53, KeyModifiers.Control | KeyModifiers.Shift, Keys.F7);

            hotkeyStartRecording.HotKeyPressed += HotkeyStartRecording_Pressed;
            hotkeyStopRecording.HotKeyPressed += HotkeyStopRecording_Pressed;
            hotkeyPlayRecording.HotKeyPressed += HotkeyPlayRecording_Pressed;
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            hotkeyStartRecording?.Dispose();
            hotkeyStartRecording = null;
            hotkeyStopRecording?.Dispose();
            hotkeyStopRecording = null;
            hotkeyPlayRecording?.Dispose();
            hotkeyPlayRecording = null;
        }
    }
}
