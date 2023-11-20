using AutoIt;
using Silence.Simulation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Forms;

namespace Silence.Macro
{

    /// <summary>
    /// Provides support for playback of recorded macros.
    /// </summary>
    public class MacroPlayer
    {

        /// <summary>
        /// Holds the mouse simulator that underlies this object.
        /// </summary>
        private MouseSimulator underlyingMouseSimulator;

        /// <summary>
        /// Holds the keyboard simulator that underlies this object.
        /// </summary>
        private KeyboardSimulator underlyingKeyboardSimulator;

        /// <summary>
        /// Holds the thread that the macro is currently executing on.
        /// </summary>
        private Thread macroThread;

        /// <summary>
        /// Holds whether or not the macro has been cancelled.
        /// </summary>
        private bool cancelled;

        /// <summary>
        /// Holds the number of times the macro should repeat before ending.
        /// </summary>
        private int repetitions;

        /// <summary>
        /// Gets the macro currently loaded into the player.
        /// </summary>
        public Macro CurrentMacro { get; private set; }

        /// <summary>
        /// Gets or sets the number of times the macro should repeat before ending.
        /// </summary>
        public int Repetitions { 
            get 
            {
                return repetitions;
            }
            set
            {
                repetitions = (value == 0 ? 1 : value);
            }
        }

        /// <summary>
        /// Gets whether or not a macro is currently being played.
        /// </summary>
        public bool IsPlaying { get; private set; }

        /// <summary>
        /// Initialies a new instance of a macro player.
        /// </summary>
        public MacroPlayer()
        {
            underlyingMouseSimulator = new MouseSimulator(new InputSimulator());
            underlyingKeyboardSimulator = new KeyboardSimulator(new InputSimulator());
            repetitions = 1;
            cancelled = false;

            AutoItX.AutoItSetOption("SendKeyDelay", 0);
            AutoItX.AutoItSetOption("SendKeyDownDelay", 0);
            AutoItX.AutoItSetOption("WinWaitDelay", 0);
        }

        /// <summary>
        /// Cancels playback of the current macro.
        /// </summary>
        public void CancelPlayback()
        {
            cancelled = IsPlaying;
        }

        /// <summary>
        /// Loads a macro into the player.
        /// </summary>
        /// <param name="macro">The macro to load into the player.</param>
        public void LoadMacro(Macro macro) 
        {
            CurrentMacro = macro;
        }

        /// <summary>
        /// Converts a point to absolute screen coordinates.
        /// </summary>
        /// <param name="point">The point to convert.</param>
        /// <returns></returns>
        private Point ConvertPointToAbsolute(Point point)
        {
            return new Point((Convert.ToDouble(65535) * point.X) / Convert.ToDouble(Screen.PrimaryScreen.Bounds.Width),
                (Convert.ToDouble(65535) * point.Y) / Convert.ToDouble(Screen.PrimaryScreen.Bounds.Height));
        }

        /// <summary>
        /// Plays the macro on the current thread.
        /// </summary>
        private void PlayMacro()
        {
            IsPlaying = true;

            // Repeat macro as required.
            for (int i = 0; i < Repetitions; i++)
            {
                // Loop through each macro event.
                foreach (MacroEvent current in CurrentMacro.Events)
                {
                    // Cancel playback.
                    if (cancelled)
                    {
                        cancelled = false;
                        i = Repetitions;
                        break;
                    }

                    // Cast event to appropriate type.
                    if(current is MacroDelayEvent) 
                    {
                        // Delay event.
                        MacroDelayEvent castEvent = (MacroDelayEvent)current;
                        Thread.Sleep(new TimeSpan(castEvent.Delay));
                    }
                    else if (current is MacroMouseMoveEvent)
                    {
                        // Mouse move event.
                        MacroMouseMoveEvent castEvent = (MacroMouseMoveEvent)current;
                        Point absolutePoint = ConvertPointToAbsolute(castEvent.Location);
                        underlyingMouseSimulator.MoveMouseTo(absolutePoint.X, absolutePoint.Y);
                    }
                    else if (current is MacroMouseDownEvent)
                    {
                        // Mouse down event.
                        MacroMouseDownEvent castEvent = (MacroMouseDownEvent)current;
                        Point absolutePoint = ConvertPointToAbsolute(castEvent.Location);
                        underlyingMouseSimulator.MoveMouseTo(absolutePoint.X, absolutePoint.Y);
                        if (castEvent.Button == System.Windows.Input.MouseButton.Left)
                        {
                            underlyingMouseSimulator.LeftButtonDown();
                        }
                        else if (castEvent.Button == System.Windows.Input.MouseButton.Right)
                        {
                            underlyingMouseSimulator.RightButtonDown();
                        }
                    }
                    else if (current is MacroMouseUpEvent)
                    {
                        // Mouse up event.
                        MacroMouseUpEvent castEvent = (MacroMouseUpEvent)current;
                        Point absolutePoint = ConvertPointToAbsolute(castEvent.Location);
                        underlyingMouseSimulator.MoveMouseTo(absolutePoint.X, absolutePoint.Y);
                        if (castEvent.Button == System.Windows.Input.MouseButton.Left)
                        {
                            underlyingMouseSimulator.LeftButtonUp();
                        }
                        else if (castEvent.Button == System.Windows.Input.MouseButton.Right)
                        {
                            underlyingMouseSimulator.RightButtonUp();
                        }
                    }
                    else if (current is MacroKeyDownEvent)
                    {
                        /*
                        // Key down event.
                        MacroKeyDownEvent castEvent = (MacroKeyDownEvent)current;
                        underlyingKeyboardSimulator.KeyDown((Simulation.Native.VirtualKeyCode)castEvent.VirtualKeyCode);
                        */
                        MacroKeyEvent castEvent = (MacroKeyEvent)current;
                        AutoItX.Send("{" + ConvertKeyToAutoIt(castEvent.VirtualKeyCode) + " down}");
                    }
                    else if (current is MacroKeyUpEvent)
                    {
                        /*
                        // Key up event.
                        MacroKeyUpEvent castEvent = (MacroKeyUpEvent)current;
                        underlyingKeyboardSimulator.KeyUp((Simulation.Native.VirtualKeyCode)castEvent.VirtualKeyCode);
                        */
                        MacroKeyEvent castEvent = (MacroKeyEvent)current;
                        AutoItX.Send("{" + ConvertKeyToAutoIt(castEvent.VirtualKeyCode) + " up}");
                    }
                    else if (current is MacroMouseWheelEvent)
                    {
                        // Mouse wheel event.
                        MacroMouseWheelEvent castEvent = (MacroMouseWheelEvent)current;
                        Point absolutePoint = ConvertPointToAbsolute(castEvent.Location);
                        underlyingMouseSimulator.MoveMouseTo(absolutePoint.X, absolutePoint.Y);
                        underlyingMouseSimulator.VerticalScroll(castEvent.Delta / 120);
                    }

                }
            }

            IsPlaying = false;
        }

        private string ConvertKeyToAutoIt(int virtualKeyCode)
        {
            var nCode = (Simulation.Native.VirtualKeyCode)virtualKeyCode;

            switch (nCode)
            {
                case Simulation.Native.VirtualKeyCode.RETURN:
                    return "ENTER";
                case Simulation.Native.VirtualKeyCode.LCONTROL:
                case Simulation.Native.VirtualKeyCode.CONTROL:
                    return "LCTRL";
                case Simulation.Native.VirtualKeyCode.RCONTROL:
                    return "RCTRL";
                case Simulation.Native.VirtualKeyCode.LMENU:
                case Simulation.Native.VirtualKeyCode.MENU:
                    return "LALT";
                case Simulation.Native.VirtualKeyCode.RMENU:
                    return "RALT";
                case Simulation.Native.VirtualKeyCode.SHIFT:
                    return "LSHIFT";
                case Simulation.Native.VirtualKeyCode.PRIOR:
                    return "PGUP";
                case Simulation.Native.VirtualKeyCode.NEXT:
                    return "PGDN";
                default:
                    return Enum.GetName(typeof(Simulation.Native.VirtualKeyCode), nCode)
                        .Replace("VK_", "");
            }
        }

        /// <summary>
        /// Plays the macro on a seperate background thread.
        /// </summary>
        public void PlayMacroAsync()
        {
            macroThread = new Thread(PlayMacro);
            macroThread.Start();
        }

    }

}
