using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace MacroAutomatorGUI
{
    // Input simulation constants and structures
    public static class InputSimulator
    {
        // Constants for mouse_event
        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;
        public const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        public const int MOUSEEVENTF_RIGHTUP = 0x10;
        public const int MOUSEEVENTF_MIDDLEDOWN = 0x20;
        public const int MOUSEEVENTF_MIDDLEUP = 0x40;
        public const int MOUSEEVENTF_ABSOLUTE = 0x8000;
        public const int MOUSEEVENTF_MOVE = 0x0001;

        // Constants for keybd_event
        public const int KEYEVENTF_EXTENDEDKEY = 0x0001;
        public const int KEYEVENTF_KEYUP = 0x0002;

        // Structure for keyboard input
        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public int type;
            public InputUnion u;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;
            [FieldOffset(0)]
            public KEYBDINPUT ki;
            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT
        {
            public short wVk;
            public short wScan;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct HARDWAREINPUT
        {
            public int uMsg;
            public short wParamL;
            public short wParamH;
        }

        // Windows API functions
        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        public static extern short VkKeyScan(char ch);

        // Helper methods for input simulation
        public static void SimulateMouseClick(int x, int y, string button = "left")
        {
            // First move the cursor to the position
            SetCursorPos(x, y);
            
            // Small delay to ensure cursor has moved
            Thread.Sleep(100);
            
            // Perform click based on button type
            switch (button.ToLower())
            {
                case "left":
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                    Thread.Sleep(100); // Longer delay between down and up
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                    break;
                case "right":
                    mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                    Thread.Sleep(100);
                    mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                    break;
                case "middle":
                    mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0);
                    Thread.Sleep(100);
                    mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0);
                    break;
            }
            
            // Additional delay after click
            Thread.Sleep(100);
        }

        public static void SimulateMouseMove(int x, int y)
        {
            // First try the SetCursorPos method
            bool success = SetCursorPos(x, y);
            
            // If that fails, try using SendInput
            if (!success)
            {
                INPUT input = new INPUT();
                input.type = 0; // INPUT_MOUSE
                
                // Set absolute positioning
                input.u.mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE;
                
                // Convert to normalized coordinates (0-65535)
                input.u.mi.dx = (x * 65535) / System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;
                input.u.mi.dy = (y * 65535) / System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height;
                
                SendInput(1, new INPUT[] { input }, Marshal.SizeOf(typeof(INPUT)));
            }
            
            // Small delay to ensure cursor has moved
            Thread.Sleep(20);
        }

        public static void SimulateKeyPress(string key)
        {
            // Convert key string to virtual key code
            byte vk = GetVirtualKeyCode(key);
            
            // Send key down and up events
            keybd_event(vk, 0, 0, 0);
            Thread.Sleep(10); // Small delay between down and up
            keybd_event(vk, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void SimulateTextTyping(string text)
        {
            foreach (char c in text)
            {
                // Convert character to virtual key code
                short vk = VkKeyScan(c);
                byte virtualKey = (byte)(vk & 0xff);
                bool shift = (vk & 0x100) != 0;

                // If shift is needed, press shift key
                if (shift)
                {
                    keybd_event(0x10, 0, 0, 0); // SHIFT key down
                }

                // Press and release the key
                keybd_event(virtualKey, 0, 0, 0);
                Thread.Sleep(5);
                keybd_event(virtualKey, 0, KEYEVENTF_KEYUP, 0);
                Thread.Sleep(5);

                // If shift was pressed, release it
                if (shift)
                {
                    keybd_event(0x10, 0, KEYEVENTF_KEYUP, 0); // SHIFT key up
                }

                // Small delay between characters
                Thread.Sleep(10);
            }
        }

        // Alternative mouse click method using Windows Forms
        public static void SimulateMouseClickAlternative(int x, int y, string button = "left")
        {
            // Move cursor to position
            SetCursorPos(x, y);
            
            // Small delay to ensure cursor has moved
            Thread.Sleep(200);
            
            // Use SendKeys to simulate mouse clicks
            switch (button.ToLower())
            {
                case "left":
                    // Send a left mouse click using SendKeys
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    break;
                case "right":
                    // For right click, we can try to send the context menu key
                    System.Windows.Forms.SendKeys.SendWait("+{F10}");
                    break;
                case "middle":
                    // Middle click doesn't have a direct SendKeys equivalent
                    // We'll use the original method as fallback
                    mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0);
                    Thread.Sleep(100);
                    mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0);
                    break;
            }
            
            // Additional delay after click
            Thread.Sleep(200);
        }

        // Direct mouse click method using SendInput with absolute coordinates
        public static void SimulateMouseClickDirect(int x, int y, string button = "left")
        {
            // Convert to normalized coordinates (0-65535)
            int normalizedX = (x * 65535) / System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;
            int normalizedY = (y * 65535) / System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height;
            
            // Create INPUT structures for mouse move and clicks
            INPUT[] inputs = new INPUT[3];
            
            // Initialize all structures
            for (int i = 0; i < inputs.Length; i++)
            {
                inputs[i].type = 0; // INPUT_MOUSE
            }
            
            // First input: Move to position (absolute coordinates)
            inputs[0].u.mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE;
            inputs[0].u.mi.dx = normalizedX;
            inputs[0].u.mi.dy = normalizedY;
            
            // Second input: Mouse button down
            // Third input: Mouse button up
            switch (button.ToLower())
            {
                case "left":
                    inputs[1].u.mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
                    inputs[2].u.mi.dwFlags = MOUSEEVENTF_LEFTUP;
                    break;
                case "right":
                    inputs[1].u.mi.dwFlags = MOUSEEVENTF_RIGHTDOWN;
                    inputs[2].u.mi.dwFlags = MOUSEEVENTF_RIGHTUP;
                    break;
                case "middle":
                    inputs[1].u.mi.dwFlags = MOUSEEVENTF_MIDDLEDOWN;
                    inputs[2].u.mi.dwFlags = MOUSEEVENTF_MIDDLEUP;
                    break;
            }
            
            // Send all inputs in one go
            SendInput(3, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        // Comprehensive mouse click method that tries multiple techniques
        public static void SimulateMouseClickComprehensive(int x, int y, string button = "left")
        {
            // First try direct method
            try
            {
                Console.WriteLine($"Trying SimulateMouseClickDirect at ({x}, {y})");
                SimulateMouseClickDirect(x, y, button);
                Thread.Sleep(50); // Small delay to let the click register
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SimulateMouseClickDirect failed: {ex.Message}");
                
                // Try the original method
                try
                {
                    Console.WriteLine($"Trying SimulateMouseClick at ({x}, {y})");
                    SimulateMouseClick(x, y, button);
                    Thread.Sleep(50);
                }
                catch (Exception ex2)
                {
                    Console.WriteLine($"SimulateMouseClick failed: {ex2.Message}");
                    
                    // Try the alternative method
                    try
                    {
                        Console.WriteLine($"Trying SimulateMouseClickAlternative at ({x}, {y})");
                        SimulateMouseClickAlternative(x, y, button);
                    }
                    catch (Exception ex3)
                    {
                        Console.WriteLine($"All mouse click methods failed: {ex3.Message}");
                    }
                }
            }
        }

        private static byte GetVirtualKeyCode(string keyName)
        {
            // Handle special keys
            switch (keyName.ToLower())
            {
                case "enter": return 0x0D;
                case "tab": return 0x09;
                case "space": return 0x20;
                case "backspace": return 0x08;
                case "escape": return 0x1B;
                case "up": return 0x26;
                case "down": return 0x28;
                case "left": return 0x25;
                case "right": return 0x27;
                case "ctrl": return 0x11;
                case "alt": return 0x12;
                case "shift": return 0x10;
                case "win": return 0x5B;
                case "f1": return 0x70;
                case "f2": return 0x71;
                case "f3": return 0x72;
                case "f4": return 0x73;
                case "f5": return 0x74;
                case "f6": return 0x75;
                case "f7": return 0x76;
                case "f8": return 0x77;
                case "f9": return 0x78;
                case "f10": return 0x79;
                case "f11": return 0x7A;
                case "f12": return 0x7B;
                default:
                    // If it's a single character, convert it
                    if (keyName.Length == 1)
                    {
                        return (byte)VkKeyScan(keyName[0]);
                    }
                    return 0;
            }
        }
    }

    public enum ActionType
    {
        MouseClick,
        MouseMove,
        KeyPress,
        TypeText,
        Wait,
        AttachToWindow,
        Sleep,
        StartRecording,
        StopRecording
    }

    public class MacroSequence
    {
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public List<MacroAction> Actions { get; set; } = new List<MacroAction>();
        public double IterationDelay { get; set; } = 1.0;
        public bool IsRecording { get; private set; } = false;
        
        // Window attachment properties
        public IntPtr AttachedWindowHandle { get; set; } = IntPtr.Zero;
        public string AttachedWindowTitle { get; set; } = string.Empty;

        public MacroSequence()
        {
            Name = "New Sequence";
            Description = "A new macro sequence";
            Actions = new List<MacroAction>();
        }

        public MacroSequence(string name)
        {
            Name = name;
            Description = $"Sequence: {name}";
            Actions = new List<MacroAction>();
        }

        public void StartRecording()
        {
            IsRecording = true;
            // Clear existing actions if we're starting a new recording
            Actions.Clear();
        }

        public void StopRecording()
        {
            IsRecording = false;
        }

        public void AddRecordedAction(MacroAction action)
        {
            if (IsRecording)
            {
                Actions.Add(action);
            }
        }
        
        public void AddAction(MacroAction action)
        {
            Actions.Add(action);
        }
    }

    public class MacroAction
    {
        public ActionType Type { get; set; }
        public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();

        public MacroAction()
        {
            Parameters = new Dictionary<string, object>();
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(Type.ToString());
            sb.Append(": ");

            switch (Type)
            {
                case ActionType.MouseClick:
                    sb.Append($"Click at ({Parameters["x"]}, {Parameters["y"]})");
                    if (Parameters.ContainsKey("button"))
                    {
                        sb.Append($" with {Parameters["button"]} button");
                    }
                    break;
                case ActionType.MouseMove:
                    sb.Append($"Move to ({Parameters["x"]}, {Parameters["y"]})");
                    break;
                case ActionType.KeyPress:
                    sb.Append($"Press {Parameters["key"]}");
                    break;
                case ActionType.TypeText:
                    sb.Append($"Type \"{Parameters["text"]}\"");
                    break;
                case ActionType.Wait:
                    sb.Append($"Wait {Parameters["seconds"]} seconds");
                    break;
                case ActionType.AttachToWindow:
                    sb.Append($"Attach to window: {Parameters["window_name"]}");
                    break;
                case ActionType.Sleep:
                    sb.Append($"Sleep {Parameters["milliseconds"]} ms");
                    break;
                case ActionType.StartRecording:
                    sb.Append("Start recording");
                    break;
                case ActionType.StopRecording:
                    sb.Append("Stop recording");
                    break;
            }

            return sb.ToString();
        }

        // Factory methods for creating different types of actions
        public static MacroAction CreateMouseClick(int x, int y, string button = "left")
        {
            return new MacroAction
            {
                Type = ActionType.MouseClick,
                Parameters = new Dictionary<string, object>
                {
                    { "x", x },
                    { "y", y },
                    { "button", button }
                }
            };
        }

        public static MacroAction CreateMouseMove(int x, int y)
        {
            return new MacroAction
            {
                Type = ActionType.MouseMove,
                Parameters = new Dictionary<string, object>
                {
                    { "x", x },
                    { "y", y }
                }
            };
        }

        public static MacroAction CreateKeyPress(string key)
        {
            return new MacroAction
            {
                Type = ActionType.KeyPress,
                Parameters = new Dictionary<string, object>
                {
                    { "key", key }
                }
            };
        }

        public static MacroAction CreateTypeText(string text)
        {
            return new MacroAction
            {
                Type = ActionType.TypeText,
                Parameters = new Dictionary<string, object>
                {
                    { "text", text }
                }
            };
        }

        public static MacroAction CreateWait(double seconds)
        {
            return new MacroAction
            {
                Type = ActionType.Wait,
                Parameters = new Dictionary<string, object>
                {
                    { "seconds", seconds }
                }
            };
        }

        public static MacroAction CreateAttachToWindow(string windowTitle)
        {
            return new MacroAction
            {
                Type = ActionType.AttachToWindow,
                Parameters = new Dictionary<string, object>
                {
                    { "window_name", windowTitle }
                }
            };
        }

        public static MacroAction CreateSleep(int milliseconds)
        {
            return new MacroAction
            {
                Type = ActionType.Sleep,
                Parameters = new Dictionary<string, object>
                {
                    { "milliseconds", milliseconds }
                }
            };
        }

        public static MacroAction CreateStartRecording()
        {
            return new MacroAction
            {
                Type = ActionType.StartRecording
            };
        }

        public static MacroAction CreateStopRecording()
        {
            return new MacroAction
            {
                Type = ActionType.StopRecording
            };
        }
    }

    // Helper class for window handling
    public static class WindowHelper
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public static IntPtr FindWindowByTitle(string title)
        {
            return FindWindow(null, title);
        }

        public static bool ActivateWindow(IntPtr handle)
        {
            return SetForegroundWindow(handle);
        }

        public static bool ActivateWindowByTitle(string title)
        {
            IntPtr handle = FindWindowByTitle(title);
            if (handle != IntPtr.Zero)
            {
                return ActivateWindow(handle);
            }
            return false;
        }

        public static RECT GetWindowPosition(IntPtr handle)
        {
            RECT rect;
            GetWindowRect(handle, out rect);
            return rect;
        }
    }

    // Input recorder for capturing keyboard and mouse events
    public class InputRecorder
    {
        private MacroSequence sequence;
        private bool isRecording = false;
        private DateTime lastActionTime;

        // Hook handles
        private IntPtr mouseHook = IntPtr.Zero;
        private IntPtr keyboardHook = IntPtr.Zero;

        // Native methods for hooking
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        // Hook constants
        private const int WH_MOUSE_LL = 14;
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_MOUSEMOVE = 0x0200;

        // Hook callback delegate
        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
        private HookProc mouseProc;
        private HookProc keyboardProc;

        public InputRecorder(MacroSequence sequence)
        {
            this.sequence = sequence;
            mouseProc = MouseHookCallback;
            keyboardProc = KeyboardHookCallback;
        }

        public void StartRecording()
        {
            if (!isRecording)
            {
                isRecording = true;
                lastActionTime = DateTime.Now;

                // Set up hooks
                mouseHook = SetWindowsHookEx(WH_MOUSE_LL, mouseProc, GetModuleHandle(null), 0);
                keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, keyboardProc, GetModuleHandle(null), 0);

                sequence.StartRecording();
            }
        }

        public void StopRecording()
        {
            if (isRecording)
            {
                isRecording = false;

                // Remove hooks
                if (mouseHook != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(mouseHook);
                    mouseHook = IntPtr.Zero;
                }

                if (keyboardHook != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(keyboardHook);
                    keyboardHook = IntPtr.Zero;
                }

                sequence.StopRecording();
            }
        }

        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && isRecording)
            {
                // Add delay action based on time since last action
                AddDelayIfNeeded();

                // Get mouse coordinates
                MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                
                if (wParam.ToInt32() == WM_LBUTTONDOWN)
                {
                    // Left mouse button click
                    var clickAction = MacroAction.CreateMouseClick((int)hookStruct.pt.x, (int)hookStruct.pt.y);
                    clickAction.Parameters["button"] = "left";
                    sequence.AddRecordedAction(clickAction);
                    // lastActionTime is now updated in AddDelayIfNeeded
                }
                else if (wParam.ToInt32() == WM_RBUTTONDOWN)
                {
                    // Right mouse button click
                    var clickAction = MacroAction.CreateMouseClick((int)hookStruct.pt.x, (int)hookStruct.pt.y);
                    clickAction.Parameters["button"] = "right";
                    sequence.AddRecordedAction(clickAction);
                    // lastActionTime is now updated in AddDelayIfNeeded
                }
                else if (wParam.ToInt32() == WM_MOUSEMOVE)
                {
                    // Mouse move - only record if significant movement (every 100ms)
                    if ((DateTime.Now - lastActionTime).TotalMilliseconds > 100)
                    {
                        sequence.AddRecordedAction(MacroAction.CreateMouseMove((int)hookStruct.pt.x, (int)hookStruct.pt.y));
                        // lastActionTime is now updated in AddDelayIfNeeded
                    }
                }
            }
            return CallNextHookEx(mouseHook, nCode, wParam, lParam);
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && isRecording && wParam.ToInt32() == WM_KEYDOWN)
            {
                // Add delay action based on time since last action
                AddDelayIfNeeded();

                // Get key information
                KBDLLHOOKSTRUCT hookStruct = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                Keys key = (Keys)hookStruct.vkCode;
                
                // Record key press
                sequence.AddRecordedAction(MacroAction.CreateKeyPress(key.ToString().ToLower()));
                // lastActionTime is now updated in AddDelayIfNeeded
            }
            return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
        }

        private void AddDelayIfNeeded()
        {
            // Add a delay action if significant time has passed since the last action
            double elapsed = (DateTime.Now - lastActionTime).TotalMilliseconds;
            
            // More granular delay thresholds
            if (elapsed > 2000) // More than 2 seconds
            {
                // For longer delays, use seconds for better readability
                double seconds = Math.Round(elapsed / 1000.0, 1); // Round to 1 decimal place
                sequence.AddRecordedAction(MacroAction.CreateWait(seconds));
                Console.WriteLine($"Added delay of {seconds} seconds");
            }
            else if (elapsed > 300) // Between 300ms and 2000ms
            {
                // For medium delays, use milliseconds for precision
                sequence.AddRecordedAction(MacroAction.CreateSleep((int)elapsed));
                Console.WriteLine($"Added delay of {elapsed}ms");
            }
            // For delays under 300ms, don't add any delay action as they're likely part of the same logical action
            
            lastActionTime = DateTime.Now;
        }

        // Structures for mouse and keyboard hooks
        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }
    }
}
