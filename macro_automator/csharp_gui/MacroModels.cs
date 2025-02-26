using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Newtonsoft.Json;

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

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(Point point);
        
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

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

        public static void SimulateKeyPress(string keyCombo)
        {
            // Parse key combination (e.g., "ctrl+shift+p")
            string[] keys = keyCombo.ToLower().Split('+');
            
            // Check if this is a Ctrl+Alt combination
            bool hasCtrl = keys.Contains("ctrl") || keys.Contains("control");
            bool hasAlt = keys.Contains("alt");
            
            // For Ctrl+Alt combinations, use SendKeys which works better across applications
            if (hasCtrl && hasAlt)
            {
                try
                {
                    // Convert to SendKeys format
                    string sendKeysFormat = "";
                    
                    // Add modifiers
                    if (hasCtrl) sendKeysFormat += "^";
                    if (hasAlt) sendKeysFormat += "%";
                    if (keys.Contains("shift")) sendKeysFormat += "+";
                    
                    // Add the main key
                    string mainKeyStr = keys.LastOrDefault(k => k != "ctrl" && k != "control" && k != "alt" && k != "shift");
                    if (!string.IsNullOrEmpty(mainKeyStr))
                    {
                        // Handle special keys
                        switch (mainKeyStr)
                        {
                            case "f1": case "f2": case "f3": case "f4": case "f5": 
                            case "f6": case "f7": case "f8": case "f9": case "f10": 
                            case "f11": case "f12":
                                sendKeysFormat += "{" + mainKeyStr.ToUpper() + "}";
                                break;
                            default:
                                // For numbers and letters, just add them directly
                                sendKeysFormat += mainKeyStr;
                                break;
                        }
                    }
                    
                    Console.WriteLine($"Sending key combination {keyCombo} using SendKeys: {sendKeysFormat}");
                    
                    // Send the keys
                    SendKeys.SendWait(sendKeysFormat);
                    
                    // Add a small delay after key press
                    Thread.Sleep(100);
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error using SendKeys: {ex.Message}");
                    // Fall through to the standard method
                }
            }
            
            // Standard method for non-Ctrl+Alt combinations
            List<byte> modifierKeys = new List<byte>();
            byte mainKey = 0;
            
            foreach (string key in keys)
            {
                string trimmedKey = key.Trim();
                
                // Check if it's a modifier key
                if (trimmedKey == "ctrl" || trimmedKey == "control")
                {
                    modifierKeys.Add(0x11); // VK_CONTROL
                }
                else if (trimmedKey == "alt")
                {
                    modifierKeys.Add(0x12); // VK_MENU
                }
                else if (trimmedKey == "shift")
                {
                    modifierKeys.Add(0x10); // VK_SHIFT
                }
                else if (trimmedKey == "win" || trimmedKey == "windows")
                {
                    modifierKeys.Add(0x5B); // VK_LWIN
                }
                else
                {
                    // This is the main key
                    mainKey = GetVirtualKeyCode(trimmedKey);
                }
            }
            
            try
            {
                // Special handling for Ctrl+Alt combinations (which can be interpreted as AltGr)
                bool hasCtrlAlt = modifierKeys.Contains(0x11) && modifierKeys.Contains(0x12);
                
                // Use SendInput for more reliable key presses
                List<INPUT> inputs = new List<INPUT>();
                
                // Press all modifier keys
                foreach (byte modKey in modifierKeys)
                {
                    INPUT modKeyDown = new INPUT();
                    modKeyDown.type = 1; // INPUT_KEYBOARD
                    modKeyDown.u.ki.wVk = modKey;
                    modKeyDown.u.ki.dwFlags = 0; // Key down
                    
                    // For Ctrl+Alt combinations, use KEYEVENTF_EXTENDEDKEY for Alt
                    if (hasCtrlAlt && modKey == 0x12)
                    {
                        modKeyDown.u.ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
                    }
                    
                    inputs.Add(modKeyDown);
                }
                
                // Press the main key
                if (mainKey != 0)
                {
                    INPUT keyDown = new INPUT();
                    keyDown.type = 1; // INPUT_KEYBOARD
                    keyDown.u.ki.wVk = mainKey;
                    keyDown.u.ki.dwFlags = 0; // Key down
                    
                    // For Ctrl+Alt combinations, use KEYEVENTF_EXTENDEDKEY for the main key
                    if (hasCtrlAlt)
                    {
                        keyDown.u.ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
                    }
                    
                    inputs.Add(keyDown);
                    
                    // Release the main key
                    INPUT keyUp = new INPUT();
                    keyUp.type = 1; // INPUT_KEYBOARD
                    keyUp.u.ki.wVk = mainKey;
                    keyUp.u.ki.dwFlags = KEYEVENTF_KEYUP;
                    
                    // For Ctrl+Alt combinations, use KEYEVENTF_EXTENDEDKEY for the main key
                    if (hasCtrlAlt)
                    {
                        keyUp.u.ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
                    }
                    
                    inputs.Add(keyUp);
                }
                
                // Release all modifier keys in reverse order
                for (int i = modifierKeys.Count - 1; i >= 0; i--)
                {
                    INPUT modKeyUp = new INPUT();
                    modKeyUp.type = 1; // INPUT_KEYBOARD
                    modKeyUp.u.ki.wVk = modifierKeys[i];
                    modKeyUp.u.ki.dwFlags = KEYEVENTF_KEYUP;
                    
                    // For Ctrl+Alt combinations, use KEYEVENTF_EXTENDEDKEY for Alt
                    if (hasCtrlAlt && modifierKeys[i] == 0x12)
                    {
                        modKeyUp.u.ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
                    }
                    
                    inputs.Add(modKeyUp);
                }
                
                // Send all inputs at once for more reliable key presses
                if (inputs.Count > 0)
                {
                    SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf(typeof(INPUT)));
                    
                    // Add a small delay after key press
                    Thread.Sleep(50);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error simulating key press: {ex.Message}");
                
                // Fallback to the old method if SendInput fails
                try
                {
                    // Special handling for Ctrl+Alt combinations
                    bool hasCtrlAlt = modifierKeys.Contains(0x11) && modifierKeys.Contains(0x12);
                    int extendedFlag = hasCtrlAlt ? KEYEVENTF_EXTENDEDKEY : 0;
                    
                    // Press all modifier keys
                    foreach (byte modKey in modifierKeys)
                    {
                        uint flags = 0;
                        if (hasCtrlAlt && modKey == 0x12)
                        {
                            flags |= KEYEVENTF_EXTENDEDKEY;
                        }
                        
                        keybd_event(modKey, 0, (int)flags, 0);
                        Thread.Sleep(10);
                    }
                    
                    // Press and release the main key
                    if (mainKey != 0)
                    {
                        keybd_event(mainKey, 0, (int)extendedFlag, 0);
                        Thread.Sleep(10);
                        keybd_event(mainKey, 0, (int)(KEYEVENTF_KEYUP | extendedFlag), 0);
                        Thread.Sleep(10);
                    }
                    
                    // Release all modifier keys in reverse order
                    for (int i = modifierKeys.Count - 1; i >= 0; i--)
                    {
                        uint flags = KEYEVENTF_KEYUP;
                        if (hasCtrlAlt && modifierKeys[i] == 0x12)
                        {
                            flags |= KEYEVENTF_EXTENDEDKEY;
                        }
                        
                        keybd_event(modifierKeys[i], 0, (int)flags, 0);
                        Thread.Sleep(10);
                    }
                }
                catch (Exception fallbackEx)
                {
                    Console.WriteLine($"Fallback key press also failed: {fallbackEx.Message}");
                }
            }
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
                // Common keys
                case "enter": return 0x0D;
                case "return": return 0x0D;
                case "tab": return 0x09;
                case "space": return 0x20;
                case "spacebar": return 0x20;
                case "backspace": return 0x08;
                case "back": return 0x08;
                case "escape": return 0x1B;
                case "esc": return 0x1B;
                
                // Arrow keys
                case "up": return 0x26;
                case "down": return 0x28;
                case "left": return 0x25;
                case "right": return 0x27;
                
                // Modifier keys
                case "ctrl": return 0x11;
                case "control": return 0x11;
                case "alt": return 0x12;
                case "shift": return 0x10;
                case "win": return 0x5B;
                case "windows": return 0x5B;
                
                // Function keys
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
                
                // Navigation keys
                case "home": return 0x24;
                case "end": return 0x23;
                case "pageup": return 0x21;
                case "pagedown": return 0x22;
                case "insert": return 0x2D;
                case "ins": return 0x2D;
                case "delete": return 0x2E;
                case "del": return 0x2E;
                
                // Lock keys
                case "capslock": return 0x14;
                case "numlock": return 0x90;
                case "scrolllock": return 0x91;
                
                // Numpad keys
                case "numpad0": return 0x60;
                case "numpad1": return 0x61;
                case "numpad2": return 0x62;
                case "numpad3": return 0x63;
                case "numpad4": return 0x64;
                case "numpad5": return 0x65;
                case "numpad6": return 0x66;
                case "numpad7": return 0x67;
                case "numpad8": return 0x68;
                case "numpad9": return 0x69;
                case "multiply": return 0x6A;
                case "add": return 0x6B;
                case "subtract": return 0x6D;
                case "decimal": return 0x6E;
                case "divide": return 0x6F;
                
                // Browser/Media keys
                case "volumemute": return 0xAD;
                case "volumedown": return 0xAE;
                case "volumeup": return 0xAF;
                case "medianexttrack": return 0xB0;
                case "mediaprevtrack": return 0xB1;
                case "mediastop": return 0xB2;
                case "mediaplaypause": return 0xB3;
                
                default:
                    // If it's a single character, convert it
                    if (keyName.Length == 1)
                    {
                        short vkKeyScan = VkKeyScan(keyName[0]);
                        if (vkKeyScan != -1)
                        {
                            return (byte)(vkKeyScan & 0xff);
                        }
                    }
                    
                    // Try to parse as a hexadecimal virtual key code
                    if (keyName.StartsWith("0x") && byte.TryParse(keyName.Substring(2), 
                        System.Globalization.NumberStyles.HexNumber, 
                        System.Globalization.CultureInfo.InvariantCulture, 
                        out byte vkCode))
                    {
                        return vkCode;
                    }
                    
                    Console.WriteLine($"Unknown key: {keyName}");
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
        public int LoopCount { get; set; } = 1;
        
        [JsonIgnore]
        public bool IsRecording { get; private set; } = false;
        
        // Window attachment properties
        [JsonIgnore]
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

        /// <summary>
        /// Creates a key press action. Supports single keys or key combinations.
        /// For combinations, use the format "ctrl+shift+p" or "alt+f4".
        /// </summary>
        /// <param name="key">Single key or key combination (e.g., "enter", "ctrl+c", "alt+shift+f4")</param>
        /// <returns>A MacroAction for pressing the specified key or key combination</returns>
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
        
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, IntPtr windowTitle);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowTextLength(IntPtr hWnd);
        
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(Point point);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        // Delegate for EnumWindows callback
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

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
        
        public static IntPtr FindWindowByPartialTitle(string partialTitle)
        {
            IntPtr foundWindow = IntPtr.Zero;
            
            // First try exact match
            foundWindow = FindWindowByTitle(partialTitle);
            if (foundWindow != IntPtr.Zero)
            {
                return foundWindow;
            }
            
            // If exact match fails, try partial match
            EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
            {
                // Only check visible windows
                if (!IsWindowVisible(hWnd))
                    return true;
                
                // Get the window title
                int length = GetWindowTextLength(hWnd);
                if (length == 0)
                    return true;
                
                StringBuilder builder = new StringBuilder(length + 1);
                GetWindowText(hWnd, builder, builder.Capacity);
                string windowTitle = builder.ToString();
                
                // Check if the window title contains our search string (case insensitive)
                if (windowTitle.IndexOf(partialTitle, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    foundWindow = hWnd;
                    return false; // Stop enumeration
                }
                
                return true; // Continue enumeration
            }, IntPtr.Zero);
            
            return foundWindow;
        }

        public static bool ActivateWindow(IntPtr handle)
        {
            return SetForegroundWindow(handle);
        }

        public static bool ActivateWindowByTitle(string title)
        {
            // Try to find the window by exact or partial title
            IntPtr handle = FindWindowByPartialTitle(title);
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
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_MOUSEMOVE = 0x0200;

        // Hook callback delegate
        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        private HookProc mouseProc;
        private HookProc keyboardProc;

        // Track modifier key states
        private bool ctrlDown = false;
        private bool shiftDown = false;
        private bool altDown = false;
        private bool winDown = false;
        
        // Pending key combination
        private List<string> currentKeyCombo = new List<string>();
        private DateTime lastKeyDown = DateTime.MinValue;
        private bool processingCombo = false;

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
                
                // Check if this is a click on the Stop Recording button
                bool isStopRecordingButtonClick = false;
                
                // If this is a click that might be on the Stop Recording button, we should ignore it
                if (wParam.ToInt32() == WM_LBUTTONDOWN || wParam.ToInt32() == WM_RBUTTONDOWN)
                {
                    // Get the control at the current mouse position
                    Point mousePoint = new Point((int)hookStruct.pt.x, (int)hookStruct.pt.y);
                    IntPtr hwnd = InputSimulator.WindowFromPoint(mousePoint);
                    
                    // If the click is on a button control, it might be the Stop Recording button
                    // We'll ignore it to prevent recording the click that stops the recording
                    if (hwnd != IntPtr.Zero)
                    {
                        StringBuilder className = new StringBuilder(256);
                        InputSimulator.GetClassName(hwnd, className, className.Capacity);
                        
                        // Check if it's a button
                        if (className.ToString().ToLower().Contains("button"))
                        {
                            isStopRecordingButtonClick = true;
                        }
                    }
                }
                
                if (!isStopRecordingButtonClick)
                {
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
            }
            return CallNextHookEx(mouseHook, nCode, wParam, lParam);
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && isRecording)
            {
                // Get key information
                KBDLLHOOKSTRUCT hookStruct = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                Keys key = (Keys)hookStruct.vkCode;
                int wParamInt = wParam.ToInt32();
                
                // Handle key down events
                if (wParamInt == WM_KEYDOWN || wParamInt == WM_SYSKEYDOWN)
                {
                    // Track modifier keys
                    if (key == Keys.ControlKey || key == Keys.LControlKey || key == Keys.RControlKey)
                    {
                        ctrlDown = true;
                        if (!processingCombo)
                        {
                            processingCombo = true;
                            currentKeyCombo.Clear();
                            currentKeyCombo.Add("ctrl");
                            lastKeyDown = DateTime.Now;
                        }
                    }
                    else if (key == Keys.ShiftKey || key == Keys.LShiftKey || key == Keys.RShiftKey)
                    {
                        shiftDown = true;
                        if (!processingCombo)
                        {
                            processingCombo = true;
                            currentKeyCombo.Clear();
                            currentKeyCombo.Add("shift");
                            lastKeyDown = DateTime.Now;
                        }
                        else if (!currentKeyCombo.Contains("shift"))
                        {
                            currentKeyCombo.Add("shift");
                        }
                    }
                    else if (key == Keys.Menu || key == Keys.LMenu || key == Keys.RMenu)
                    {
                        altDown = true;
                        if (!processingCombo)
                        {
                            processingCombo = true;
                            currentKeyCombo.Clear();
                            currentKeyCombo.Add("alt");
                            lastKeyDown = DateTime.Now;
                        }
                        else if (!currentKeyCombo.Contains("alt"))
                        {
                            currentKeyCombo.Add("alt");
                        }
                    }
                    else if (key == Keys.LWin || key == Keys.RWin)
                    {
                        winDown = true;
                        if (!processingCombo)
                        {
                            processingCombo = true;
                            currentKeyCombo.Clear();
                            currentKeyCombo.Add("win");
                            lastKeyDown = DateTime.Now;
                        }
                        else if (!currentKeyCombo.Contains("win"))
                        {
                            currentKeyCombo.Add("win");
                        }
                    }
                    else
                    {
                        // Non-modifier key
                        // Add delay action based on time since last action (only if not part of a combo)
                        if (!processingCombo)
                        {
                            AddDelayIfNeeded();
                            
                            // Simple key press (no modifiers)
                            sequence.AddRecordedAction(MacroAction.CreateKeyPress(key.ToString().ToLower()));
                        }
                        else
                        {
                            // Part of a key combination
                            string keyName = GetProperKeyName(key);
                            
                            // Add the non-modifier key to the combination
                            if (!currentKeyCombo.Contains(keyName))
                            {
                                currentKeyCombo.Add(keyName);
                            }
                            
                            // If it's been more than 50ms since we started tracking this combo,
                            // or if this is a key that likely completes a combo, record it
                            if ((DateTime.Now - lastKeyDown).TotalMilliseconds > 50 || 
                                IsLikelyComboCompletionKey(key))
                            {
                                AddDelayIfNeeded();
                                
                                // Create the key combination string (e.g., "ctrl+shift+p")
                                string keyCombo = string.Join("+", currentKeyCombo);
                                
                                // Record the key combination
                                sequence.AddRecordedAction(MacroAction.CreateKeyPress(keyCombo));
                                
                                // Reset combo tracking
                                processingCombo = false;
                                currentKeyCombo.Clear();
                            }
                        }
                    }
                }
                // Handle key up events
                else if (wParamInt == WM_KEYUP || wParamInt == WM_SYSKEYUP)
                {
                    // Update modifier key states
                    if (key == Keys.ControlKey || key == Keys.LControlKey || key == Keys.RControlKey)
                    {
                        ctrlDown = false;
                    }
                    else if (key == Keys.ShiftKey || key == Keys.LShiftKey || key == Keys.RShiftKey)
                    {
                        shiftDown = false;
                    }
                    else if (key == Keys.Menu || key == Keys.LMenu || key == Keys.RMenu)
                    {
                        altDown = false;
                    }
                    else if (key == Keys.LWin || key == Keys.RWin)
                    {
                        winDown = false;
                    }
                    
                    // If all modifier keys are up and we were processing a combo but didn't complete it,
                    // record whatever we have so far
                    if (processingCombo && !ctrlDown && !shiftDown && !altDown && !winDown && currentKeyCombo.Count > 0)
                    {
                        // If the combo only has modifier keys, don't record it
                        if (currentKeyCombo.Count > 1 || 
                            (!currentKeyCombo.Contains("ctrl") && 
                             !currentKeyCombo.Contains("shift") && 
                             !currentKeyCombo.Contains("alt") && 
                             !currentKeyCombo.Contains("win")))
                        {
                            AddDelayIfNeeded();
                            
                            // Create the key combination string
                            string keyCombo = string.Join("+", currentKeyCombo);
                            
                            // Record the key combination
                            sequence.AddRecordedAction(MacroAction.CreateKeyPress(keyCombo));
                        }
                        
                        // Reset combo tracking
                        processingCombo = false;
                        currentKeyCombo.Clear();
                    }
                }
            }
            return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
        }
        
        // Helper method to get proper key name for keyboard hook
        private string GetProperKeyName(Keys key)
        {
            // Handle numeric keys (D0-D9)
            if (key >= Keys.D0 && key <= Keys.D9)
            {
                // Convert D0-D9 to 0-9
                return (key - Keys.D0).ToString();
            }
            
            // Handle numeric keypad keys (NumPad0-NumPad9)
            if (key >= Keys.NumPad0 && key <= Keys.NumPad9)
            {
                // Convert NumPad0-NumPad9 to 0-9
                return (key - Keys.NumPad0).ToString();
            }
            
            // For all other keys, use the standard key name in lowercase
            return key.ToString().ToLower();
        }
        
        private bool IsLikelyComboCompletionKey(Keys key)
        {
            // Keys that are commonly used to complete hotkey combinations
            switch (key)
            {
                case Keys.A: // Ctrl+A (Select All)
                case Keys.C: // Ctrl+C (Copy)
                case Keys.V: // Ctrl+V (Paste)
                case Keys.X: // Ctrl+X (Cut)
                case Keys.Z: // Ctrl+Z (Undo)
                case Keys.Y: // Ctrl+Y (Redo)
                case Keys.S: // Ctrl+S (Save)
                case Keys.O: // Ctrl+O (Open)
                case Keys.P: // Ctrl+P (Print)
                case Keys.F: // Ctrl+F (Find)
                case Keys.N: // Ctrl+N (New)
                case Keys.W: // Ctrl+W (Close)
                case Keys.T: // Ctrl+T (New Tab)
                case Keys.Tab: // Ctrl+Tab, Alt+Tab
                case Keys.Escape:
                case Keys.Delete:
                case Keys.Home:
                case Keys.End:
                case Keys.PageUp:
                case Keys.PageDown:
                case Keys.Up:
                case Keys.Down:
                case Keys.Left:
                case Keys.Right:
                case Keys.Enter:
                case Keys.F1:
                case Keys.F2:
                case Keys.F3:
                case Keys.F4:
                case Keys.F5:
                case Keys.F6:
                case Keys.F7:
                case Keys.F8:
                case Keys.F9:
                case Keys.F10:
                case Keys.F11:
                case Keys.F12:
                    return true;
                default:
                    return false;
            }
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
