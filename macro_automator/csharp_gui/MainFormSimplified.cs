using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using YamlDotNet.Serialization;

namespace MacroAutomatorGUI
{
    public class MainFormSimplified : Form
    {
        private List<MacroSequence> macroSequences = new List<MacroSequence>();
        private bool isRunning = false;
        private string configPath;
        
        // UI Controls
        private ListBox sequencesList;
        private RichTextBox logTextBox;
        private Button startButton;
        private Button stopButton;
        private Button recordButton;
        private Button saveButton;
        private Button deleteButton;
        private Button loadButton;
        private NumericUpDown iterationsInput;
        private CheckBox loopForeverCheckbox;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel;

        // Recording
        private InputRecorder inputRecorder;
        private bool isRecording = false;
        
        public MainFormSimplified(string configPath = null)
        {
            this.configPath = configPath ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sequences.yaml");
            InitializeComponent();
            LoadSequences();
            SetupHotkeys();
        }

        private void InitializeComponent()
        {
            // Form settings
            this.Text = "Macro Automator";
            this.Size = new Size(700, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(30, 30, 30);
            this.ForeColor = Color.White;
            
            // Main layout
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 3,
                ColumnCount = 1,
                BackColor = Color.FromArgb(30, 30, 30)
            };
            
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 120));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 150));
            
            this.Controls.Add(mainLayout);

            // Top panel (controls)
            Panel controlPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(40, 40, 40)
            };
            
            mainLayout.Controls.Add(controlPanel, 0, 0);

            // Control buttons
            startButton = new Button
            {
                Text = "Start Sequence (F5)",
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(150, 40),
                Location = new Point(20, 20)
            };
            startButton.Click += StartButton_Click;
            controlPanel.Controls.Add(startButton);

            stopButton = new Button
            {
                Text = "Stop (F6)",
                BackColor = Color.FromArgb(200, 0, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(150, 40),
                Location = new Point(180, 20)
            };
            stopButton.Click += StopButton_Click;
            controlPanel.Controls.Add(stopButton);
            
            recordButton = new Button
            {
                Text = "Record New Sequence",
                BackColor = Color.FromArgb(180, 0, 180),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(150, 40),
                Location = new Point(340, 20)
            };
            recordButton.Click += RecordButton_Click;
            controlPanel.Controls.Add(recordButton);

            loadButton = new Button
            {
                Text = "Load Sequence",
                BackColor = Color.FromArgb(0, 120, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(150, 40),
                Location = new Point(500, 20)
            };
            loadButton.Click += LoadButton_Click;
            controlPanel.Controls.Add(loadButton);

            Label iterationsLabel = new Label
            {
                Text = "Iterations:",
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(20, 75)
            };
            controlPanel.Controls.Add(iterationsLabel);

            iterationsInput = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 9999,
                Value = 1,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                Size = new Size(80, 25),
                Location = new Point(90, 73)
            };
            controlPanel.Controls.Add(iterationsInput);

            loopForeverCheckbox = new CheckBox
            {
                Text = "Loop Forever",
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(180, 75)
            };
            controlPanel.Controls.Add(loopForeverCheckbox);

            // Middle panel (sequences)
            Panel sequencesPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(35, 35, 35),
                Padding = new Padding(10)
            };
            
            mainLayout.Controls.Add(sequencesPanel, 0, 1);

            TableLayoutPanel sequencesLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 3,
                ColumnCount = 1,
                BackColor = Color.FromArgb(35, 35, 35)
            };
            
            sequencesLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30)); // Header
            sequencesLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // List
            sequencesLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50)); // Buttons
            sequencesPanel.Controls.Add(sequencesLayout);

            Label sequencesLabel = new Label
            {
                Text = "Available Sequences",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                AutoSize = true
            };
            sequencesLayout.Controls.Add(sequencesLabel, 0, 0);

            sequencesList = new ListBox
            {
                BackColor = Color.FromArgb(50, 50, 50),
                ForeColor = Color.White,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F)
            };
            sequencesList.SelectedIndexChanged += SequencesList_SelectedIndexChanged;
            sequencesLayout.Controls.Add(sequencesList, 0, 1);

            // Add buttons panel for save and delete
            Panel sequenceButtonsPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(35, 35, 35)
            };
            sequencesLayout.Controls.Add(sequenceButtonsPanel, 0, 2);

            saveButton = new Button
            {
                Text = "Save Sequence",
                BackColor = Color.FromArgb(0, 120, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(120, 30),
                Location = new Point(10, 10)
            };
            saveButton.Click += SaveButton_Click;
            sequenceButtonsPanel.Controls.Add(saveButton);

            deleteButton = new Button
            {
                Text = "Delete Sequence",
                BackColor = Color.FromArgb(120, 0, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(120, 30),
                Location = new Point(140, 10)
            };
            deleteButton.Click += DeleteButton_Click;
            sequenceButtonsPanel.Controls.Add(deleteButton);

            // Bottom panel (log)
            Panel logPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(25, 25, 25),
                Padding = new Padding(10)
            };
            
            mainLayout.Controls.Add(logPanel, 0, 2);

            TableLayoutPanel logLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = Color.FromArgb(25, 25, 25)
            };
            
            logLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            logLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            logPanel.Controls.Add(logLayout);

            Label logLabel = new Label
            {
                Text = "Log",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                AutoSize = true,
                Dock = DockStyle.Fill
            };
            logLayout.Controls.Add(logLabel, 0, 0);

            logTextBox = new RichTextBox
            {
                BackColor = Color.FromArgb(45, 45, 45),
                ForeColor = Color.LightGray,
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F)
            };
            logLayout.Controls.Add(logTextBox, 0, 1);

            // Status strip
            statusStrip = new StatusStrip
            {
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.White
            };
            
            statusLabel = new ToolStripStatusLabel("Ready");
            statusStrip.Items.Add(statusLabel);
            this.Controls.Add(statusStrip);
        }

        private void SetupHotkeys()
        {
            this.KeyPreview = true;
            this.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.F5)
                {
                    StartButton_Click(sender, e);
                }
                else if (e.KeyCode == Keys.F6)
                {
                    StopButton_Click(sender, e);
                }
            };
        }

        private void LoadSequences()
        {
            AddLogMessage("Loading sequences...");
            
            // Clear existing sequences
            macroSequences.Clear();
            
            // Create sequences directory if it doesn't exist
            string sequencesDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sequences");
            if (!Directory.Exists(sequencesDir))
            {
                Directory.CreateDirectory(sequencesDir);
                AddLogMessage("Created sequences directory");
            }
            
            // First add example sequences
            AddLogMessage("Adding example sequences.");
            AddExampleSequences();
            
            // Keep track of example sequence names to avoid duplicates
            HashSet<string> existingSequenceNames = new HashSet<string>(
                macroSequences.Select(seq => seq.Name)
            );
            
            // Load sequences from files
            try
            {
                // Load JSON files
                string[] jsonFiles = Directory.GetFiles(sequencesDir, "*.json");
                foreach (string file in jsonFiles)
                {
                    try
                    {
                        string json = File.ReadAllText(file);
                        MacroSequence sequence = JsonConvert.DeserializeObject<MacroSequence>(json);
                        if (sequence != null)
                        {
                            // Only add if not a duplicate of an example sequence
                            if (!existingSequenceNames.Contains(sequence.Name))
                            {
                                macroSequences.Add(sequence);
                                existingSequenceNames.Add(sequence.Name);
                                AddLogMessage($"Loaded JSON sequence: {sequence.Name}");
                            }
                            else
                            {
                                AddLogMessage($"Skipped duplicate sequence: {sequence.Name}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        AddLogMessage($"Error loading JSON sequence file {Path.GetFileName(file)}: {ex.Message}");
                    }
                }
                
                // Load YAML files
                string[] yamlFiles = Directory.GetFiles(sequencesDir, "*.yaml");
                string[] ymlFiles = Directory.GetFiles(sequencesDir, "*.yml");
                string[] allYamlFiles = yamlFiles.Concat(ymlFiles).ToArray();
                
                foreach (string file in allYamlFiles)
                {
                    try
                    {
                        // Load YAML file and convert to MacroSequence objects
                        var yamlSequences = LoadSequencesFromYamlFile(file);
                        
                        foreach (var sequence in yamlSequences)
                        {
                            // Only add if not a duplicate of an example sequence
                            if (!existingSequenceNames.Contains(sequence.Name))
                            {
                                macroSequences.Add(sequence);
                                existingSequenceNames.Add(sequence.Name);
                                AddLogMessage($"Loaded YAML sequence: {sequence.Name}");
                            }
                            else
                            {
                                AddLogMessage($"Skipped duplicate sequence: {sequence.Name}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        AddLogMessage($"Error loading YAML sequence file {Path.GetFileName(file)}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                AddLogMessage($"Error accessing sequences directory: {ex.Message}");
            }
            
            // Update the list
            sequencesList.Items.Clear();
            foreach (var sequence in macroSequences)
            {
                sequencesList.Items.Add(sequence.Name);
            }
            
            if (sequencesList.Items.Count > 0)
            {
                sequencesList.SelectedIndex = 0;
            }
            
            AddLogMessage($"Loaded {macroSequences.Count} sequences");
        }
        
        private List<MacroSequence> LoadSequencesFromYamlFile(string filePath)
        {
            List<MacroSequence> sequences = new List<MacroSequence>();
            
            try
            {
                // Load the YAML content
                var yamlContent = YamlHelper.LoadFromYaml<Dictionary<string, object>>(filePath);
                
                // Check if the YAML has the expected structure
                if (yamlContent.ContainsKey("macros") && yamlContent["macros"] is Dictionary<object, object> macrosDict)
                {
                    // Process each macro in the YAML file
                    foreach (var macroEntry in macrosDict)
                    {
                        string sequenceName = macroEntry.Key.ToString();
                        var macroData = macroEntry.Value as Dictionary<object, object>;
                        
                        if (macroData != null && macroData.ContainsKey("actions"))
                        {
                            // Create a new sequence
                            MacroSequence sequence = new MacroSequence
                            {
                                Name = sequenceName,
                                Description = $"Imported from {Path.GetFileName(filePath)}",
                                Actions = new List<MacroAction>()
                            };
                            
                            // Set iteration delay if present
                            if (macroData.ContainsKey("iteration_delay") && macroData["iteration_delay"] is double iterDelay)
                            {
                                sequence.IterationDelay = iterDelay;
                            }
                            
                            // Set loop count if present
                            if (macroData.ContainsKey("loop_count") && macroData["loop_count"] is int loopCount)
                            {
                                sequence.LoopCount = loopCount;
                            }
                            
                            // Process actions
                            var actionsList = macroData["actions"] as List<object>;
                            if (actionsList != null)
                            {
                                foreach (var actionObj in actionsList)
                                {
                                    var actionDict = actionObj as Dictionary<object, object>;
                                    if (actionDict != null && actionDict.ContainsKey("type"))
                                    {
                                        string actionType = actionDict["type"].ToString().ToLower();
                                        MacroAction action = ConvertYamlActionToMacroAction(actionType, actionDict);
                                        if (action != null)
                                        {
                                            sequence.Actions.Add(action);
                                        }
                                    }
                                }
                            }
                            
                            sequences.Add(sequence);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AddLogMessage($"Error parsing YAML file {Path.GetFileName(filePath)}: {ex.Message}");
            }
            
            return sequences;
        }
        
        private MacroAction ConvertYamlActionToMacroAction(string yamlActionType, Dictionary<object, object> actionData)
        {
            // Convert YAML action type to C# ActionType
            ActionType actionType;
            switch (yamlActionType)
            {
                case "mouse_click":
                    actionType = ActionType.MouseClick;
                    break;
                case "mouse_move":
                    actionType = ActionType.MouseMove;
                    break;
                case "key_press":
                    actionType = ActionType.KeyPress;
                    break;
                case "type_text":
                    actionType = ActionType.TypeText;
                    break;
                case "wait":
                    actionType = ActionType.Wait;
                    break;
                case "sleep":
                    actionType = ActionType.Sleep;
                    break;
                case "attach_to_window":
                case "find_window":
                    actionType = ActionType.AttachToWindow;
                    break;
                default:
                    AddLogMessage($"Unknown action type in YAML: {yamlActionType}");
                    return null;
            }
            
            // Create action with parameters
            MacroAction action = new MacroAction
            {
                Type = actionType,
                Parameters = new Dictionary<string, object>()
            };
            
            // Convert parameters based on action type
            foreach (var param in actionData)
            {
                string key = param.Key.ToString();
                if (key != "type") // Skip the type parameter as we've already processed it
                {
                    // Handle specific parameter mappings
                    string paramName = key;
                    
                    // Map YAML parameter names to C# parameter names
                    switch (key)
                    {
                        case "window_name":
                            if (actionType == ActionType.AttachToWindow)
                                paramName = "window_name";
                            break;
                        case "text":
                            if (actionType == ActionType.TypeText)
                                paramName = "text";
                            break;
                        case "seconds":
                            if (actionType == ActionType.Wait)
                                paramName = "seconds";
                            else if (actionType == ActionType.Sleep)
                                paramName = "milliseconds";
                            break;
                    }
                    
                    // Convert value if needed
                    object value = param.Value;
                    if (paramName == "milliseconds" && key == "seconds" && value is double seconds)
                    {
                        value = (int)(seconds * 1000); // Convert seconds to milliseconds
                    }
                    
                    action.Parameters[paramName] = value;
                }
            }
            
            return action;
        }

        private void AddExampleSequences()
        {
            // Add example sequences
            macroSequences.Add(new MacroSequence
            {
                Name = "Click Center",
                Description = "Clicks at the center of the active window",
                Actions = new List<MacroAction>
                {
                    MacroAction.CreateMouseClick(400, 300)
                }
            });
            
            macroSequences.Add(new MacroSequence
            {
                Name = "Type Hello World",
                Description = "Types 'Hello World' in the active window",
                Actions = new List<MacroAction>
                {
                    MacroAction.CreateTypeText("Hello World")
                }
            });
            
            macroSequences.Add(new MacroSequence
            {
                Name = "Press Enter 5 Times",
                Description = "Presses Enter key 5 times with delays",
                Actions = new List<MacroAction>
                {
                    MacroAction.CreateKeyPress("enter"),
                    MacroAction.CreateWait(0.5),
                    MacroAction.CreateKeyPress("enter"),
                    MacroAction.CreateWait(0.5),
                    MacroAction.CreateKeyPress("enter"),
                    MacroAction.CreateWait(0.5),
                    MacroAction.CreateKeyPress("enter"),
                    MacroAction.CreateWait(0.5),
                    MacroAction.CreateKeyPress("enter")
                }
            });
            
            // Add new example sequences with the new action types
            macroSequences.Add(new MacroSequence
            {
                Name = "Attach to Notepad",
                Description = "Attaches to Notepad window and types text",
                Actions = new List<MacroAction>
                {
                    MacroAction.CreateAttachToWindow("Notepad"),
                    MacroAction.CreateSleep(500),
                    MacroAction.CreateTypeText("Hello from Macro Automator!")
                }
            });
            
            macroSequences.Add(new MacroSequence
            {
                Name = "Sleep Example",
                Description = "Demonstrates sleep action with millisecond precision",
                Actions = new List<MacroAction>
                {
                    MacroAction.CreateTypeText("Starting sleep demo"),
                    MacroAction.CreateSleep(1000), // 1 second
                    MacroAction.CreateTypeText("After 1 second"),
                    MacroAction.CreateSleep(500),  // 0.5 seconds
                    MacroAction.CreateTypeText("After 0.5 seconds")
                }
            });
            
            // Add hotkey combination example
            macroSequences.Add(new MacroSequence
            {
                Name = "Hotkey Combinations",
                Description = "Demonstrates various hotkey combinations",
                Actions = new List<MacroAction>
                {
                    MacroAction.CreateTypeText("Demonstrating hotkey combinations:"),
                    MacroAction.CreateSleep(1000),
                    MacroAction.CreateKeyPress("ctrl+a"), // Select all
                    MacroAction.CreateSleep(500),
                    MacroAction.CreateKeyPress("ctrl+c"), // Copy
                    MacroAction.CreateSleep(500),
                    MacroAction.CreateKeyPress("right"), // Move cursor right
                    MacroAction.CreateSleep(500),
                    MacroAction.CreateKeyPress("ctrl+v"), // Paste
                    MacroAction.CreateSleep(500),
                    MacroAction.CreateKeyPress("ctrl+shift+left"), // Select word to left
                    MacroAction.CreateSleep(500),
                    MacroAction.CreateKeyPress("delete") // Delete selection
                }
            });
        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            if (isRunning)
            {
                AddLogMessage("A sequence is already running");
                return;
            }

            if (sequencesList.SelectedIndex < 0)
            {
                AddLogMessage("No sequence selected");
                return;
            }

            MacroSequence sequence = macroSequences[sequencesList.SelectedIndex];
            
            // Update the sequence's loop count based on UI settings
            if (loopForeverCheckbox.Checked)
            {
                sequence.LoopCount = 0; // 0 means infinite
            }
            else
            {
                sequence.LoopCount = (int)iterationsInput.Value;
            }
            
            int iterations = sequence.LoopCount;

            AddLogMessage($"Starting sequence: {sequence.Name}");
            AddLogMessage($"Iterations: {(iterations == 0 ? "Infinite" : iterations.ToString())}");

            isRunning = true;
            UpdateStatus($"Running: {sequence.Name}");

            // Simulate execution
            Task.Run(() =>
            {
                try
                {
                    for (int i = 0; i < (iterations == 0 ? 999999 : iterations); i++)
                    {
                        if (!isRunning) break;
                        
                        AddLogMessage($"Iteration {i + 1}");
                        
                        foreach (var action in sequence.Actions)
                        {
                            if (!isRunning) break;
                            AddLogMessage($"Executing: {action}");
                            
                            // Execute the action based on its type
                            ExecuteAction(action);
                            
                            // Small delay between actions
                            Thread.Sleep(100);
                        }
                        
                        if (iterations > 0 && i >= iterations - 1) break;
                        
                        AddLogMessage($"Iteration delay: {sequence.IterationDelay}s");
                        Thread.Sleep((int)(sequence.IterationDelay * 1000));
                    }
                }
                finally
                {
                    isRunning = false;
                    this.Invoke(new Action(() =>
                    {
                        UpdateStatus("Ready");
                        AddLogMessage("Sequence completed");
                    }));
                }
            });
        }

        private void ExecuteAction(MacroAction action)
        {
            switch (action.Type)
            {
                case ActionType.MouseClick:
                    // Implement mouse click
                    int clickX = Convert.ToInt32(action.Parameters["x"]);
                    int clickY = Convert.ToInt32(action.Parameters["y"]);
                    string button = action.Parameters.ContainsKey("button") ? action.Parameters["button"].ToString() : "left";
                    
                    AddLogMessage($"Clicking at ({clickX}, {clickY}) with {button} button");
                    
                    try
                    {
                        // Use the comprehensive method that tries multiple techniques
                        InputSimulator.SimulateMouseClickComprehensive(clickX, clickY, button);
                    }
                    catch (Exception ex)
                    {
                        AddLogMessage($"Error during mouse click: {ex.Message}");
                    }
                    break;
                    
                case ActionType.MouseMove:
                    // Implement mouse move
                    int moveX = Convert.ToInt32(action.Parameters["x"]);
                    int moveY = Convert.ToInt32(action.Parameters["y"]);
                    
                    AddLogMessage($"Moving mouse to ({moveX}, {moveY})");
                    InputSimulator.SimulateMouseMove(moveX, moveY);
                    break;
                    
                case ActionType.KeyPress:
                    // Implement key press
                    string key = action.Parameters["key"].ToString();
                    
                    AddLogMessage($"Pressing key: {key}");
                    InputSimulator.SimulateKeyPress(key);
                    break;
                    
                case ActionType.TypeText:
                    // Implement typing text
                    string text = action.Parameters["text"].ToString();
                    
                    AddLogMessage($"Typing: {text}");
                    InputSimulator.SimulateTextTyping(text);
                    break;
                    
                case ActionType.Wait:
                    // Implement wait
                    double seconds = Convert.ToDouble(action.Parameters["seconds"]);
                    AddLogMessage($"Waiting for {seconds} seconds");
                    Thread.Sleep((int)(seconds * 1000));
                    break;
                    
                case ActionType.Sleep:
                    // Implement sleep (millisecond precision)
                    int ms = Convert.ToInt32(action.Parameters["milliseconds"]);
                    AddLogMessage($"Sleeping for {ms} milliseconds");
                    Thread.Sleep(ms);
                    break;
                    
                case ActionType.AttachToWindow:
                    // Implement window attachment
                    string windowTitle = action.Parameters["window_name"].ToString();
                    AddLogMessage($"Attaching to window: {windowTitle}");
                    bool success = WindowHelper.ActivateWindowByTitle(windowTitle);
                    if (success)
                    {
                        AddLogMessage($"Successfully attached to window: {windowTitle}");
                    }
                    else
                    {
                        AddLogMessage($"Failed to find window: {windowTitle}");
                    }
                    break;
                    
                default:
                    AddLogMessage($"Unsupported action type: {action.Type}");
                    break;
            }
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            if (!isRunning)
            {
                AddLogMessage("No sequence is running");
                return;
            }

            isRunning = false;
            AddLogMessage("Stopping sequence...");
            UpdateStatus("Ready");
        }

        private void RecordButton_Click(object sender, EventArgs e)
        {
            if (isRunning)
            {
                AddLogMessage("Cannot start recording while a sequence is running");
                return;
            }
            
            if (isRecording)
            {
                // Stop recording
                StopRecording();
            }
            else
            {
                // Start recording
                StartRecording();
            }
        }
        
        private void StartRecording()
        {
            // Create a new sequence for recording
            MacroSequence newSequence = new MacroSequence($"Recorded Sequence {DateTime.Now.ToString("yyyyMMdd_HHmmss")}");
            macroSequences.Add(newSequence);
            
            // Create input recorder
            inputRecorder = new InputRecorder(newSequence);
            
            // Update UI
            recordButton.Text = "Stop Recording";
            recordButton.BackColor = Color.FromArgb(255, 0, 0);
            isRecording = true;
            
            // Show recording status in the log
            AddLogMessage("=== RECORDING INSTRUCTIONS ===");
            AddLogMessage("• Perform actions at your normal pace");
            AddLogMessage("• Delays between actions will be automatically recorded");
            AddLogMessage("• For best results, avoid unnecessary movements");
            AddLogMessage("• Press the Stop Recording button when finished");
            AddLogMessage("===============================");
            
            // Start recording
            inputRecorder.StartRecording();
            
            AddLogMessage("Recording started. Press the Stop Recording button to finish.");
            UpdateStatus("Recording...");
            
            // Update the sequences list
            sequencesList.Items.Add(newSequence.Name);
            sequencesList.SelectedIndex = sequencesList.Items.Count - 1;
            
            // Start a timer to update recording duration
            StartRecordingTimer();
        }
        
        private System.Windows.Forms.Timer recordingTimer;
        private DateTime recordingStartTime;
        
        private void StartRecordingTimer()
        {
            recordingStartTime = DateTime.Now;
            
            if (recordingTimer == null)
            {
                recordingTimer = new System.Windows.Forms.Timer();
                recordingTimer.Interval = 1000; // Update every second
                recordingTimer.Tick += RecordingTimer_Tick;
            }
            
            recordingTimer.Start();
        }
        
        private void RecordingTimer_Tick(object sender, EventArgs e)
        {
            if (isRecording)
            {
                TimeSpan duration = DateTime.Now - recordingStartTime;
                UpdateStatus($"Recording... [{duration.Minutes:00}:{duration.Seconds:00}]");
            }
        }
        
        private void StopRecording()
        {
            if (inputRecorder != null)
            {
                // Stop recording
                inputRecorder.StopRecording();
                
                // Stop the recording timer
                if (recordingTimer != null)
                {
                    recordingTimer.Stop();
                }
                
                // Calculate total recording duration
                TimeSpan recordingDuration = DateTime.Now - recordingStartTime;
                
                // Update UI
                recordButton.Text = "Record New Sequence";
                recordButton.BackColor = Color.FromArgb(180, 0, 180);
                isRecording = false;
                
                AddLogMessage("Recording stopped. The new sequence is now available in the list.");
                AddLogMessage($"Total recording time: {recordingDuration.Minutes:00}:{recordingDuration.Seconds:00}");
                UpdateStatus("Ready");
                
                // Get the recorded sequence
                MacroSequence recordedSequence = macroSequences[sequencesList.SelectedIndex];
                AddLogMessage($"Recorded {recordedSequence.Actions.Count} actions");
                
                // Show action summary
                int mouseClicks = 0;
                int keyPresses = 0;
                int delays = 0;
                
                foreach (var action in recordedSequence.Actions)
                {
                    switch (action.Type)
                    {
                        case ActionType.MouseClick:
                            mouseClicks++;
                            break;
                        case ActionType.KeyPress:
                            keyPresses++;
                            break;
                        case ActionType.Wait:
                        case ActionType.Sleep:
                            delays++;
                            break;
                    }
                }
                
                AddLogMessage($"Summary: {mouseClicks} mouse clicks, {keyPresses} key presses, {delays} delays");
            }
        }

        private void LoadButton_Click(object sender, EventArgs e)
        {
            LoadSequenceFromFile();
        }

        private void LoadSequenceFromFile()
        {
            using (OpenFileDialog openDialog = new OpenFileDialog())
            {
                string sequencesDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sequences");
                if (!Directory.Exists(sequencesDir))
                {
                    Directory.CreateDirectory(sequencesDir);
                }
                
                openDialog.InitialDirectory = sequencesDir;
                openDialog.Filter = "YAML Files (*.yaml;*.yml)|*.yaml;*.yml|JSON Files (*.json)|*.json|All Files (*.*)|*.*";
                openDialog.DefaultExt = "yaml";
                openDialog.Multiselect = false;
                
                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (openDialog.FileName.EndsWith(".json"))
                        {
                            string json = File.ReadAllText(openDialog.FileName);
                            MacroSequence sequence = JsonConvert.DeserializeObject<MacroSequence>(json);
                            
                            if (sequence != null)
                            {
                                // Check if a sequence with the same name already exists
                                int existingIndex = -1;
                                for (int i = 0; i < macroSequences.Count; i++)
                                {
                                    if (macroSequences[i].Name == sequence.Name)
                                    {
                                        existingIndex = i;
                                        break;
                                    }
                                }
                                
                                if (existingIndex >= 0)
                                {
                                    // Replace existing sequence
                                    macroSequences[existingIndex] = sequence;
                                    sequencesList.Items[existingIndex] = sequence.Name;
                                    sequencesList.SelectedIndex = existingIndex;
                                    AddLogMessage($"Updated existing sequence: {sequence.Name}");
                                }
                                else
                                {
                                    // Add new sequence
                                    macroSequences.Add(sequence);
                                    sequencesList.Items.Add(sequence.Name);
                                    sequencesList.SelectedIndex = macroSequences.Count - 1;
                                    AddLogMessage($"Loaded sequence: {sequence.Name}");
                                }
                            }
                            else
                            {
                                MessageBox.Show("Failed to load sequence. The file may be corrupted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else if (openDialog.FileName.EndsWith(".yaml") || openDialog.FileName.EndsWith(".yml"))
                        {
                            var yamlSequences = LoadSequencesFromYamlFile(openDialog.FileName);
                            foreach (var sequence in yamlSequences)
                            {
                                // Check if a sequence with the same name already exists
                                int existingIndex = -1;
                                for (int i = 0; i < macroSequences.Count; i++)
                                {
                                    if (macroSequences[i].Name == sequence.Name)
                                    {
                                        existingIndex = i;
                                        break;
                                    }
                                }
                                
                                if (existingIndex >= 0)
                                {
                                    // Replace existing sequence
                                    macroSequences[existingIndex] = sequence;
                                    sequencesList.Items[existingIndex] = sequence.Name;
                                    sequencesList.SelectedIndex = existingIndex;
                                    AddLogMessage($"Updated existing sequence: {sequence.Name}");
                                }
                                else
                                {
                                    // Add new sequence
                                    macroSequences.Add(sequence);
                                    sequencesList.Items.Add(sequence.Name);
                                    sequencesList.SelectedIndex = macroSequences.Count - 1;
                                    AddLogMessage($"Loaded sequence: {sequence.Name}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error loading sequence: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        AddLogMessage($"Error loading sequence: {ex.Message}");
                    }
                }
            }
        }

        private void SequencesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (sequencesList.SelectedIndex >= 0)
            {
                MacroSequence selectedSequence = macroSequences[sequencesList.SelectedIndex];
                
                // Update the iteration input and checkbox based on the sequence's loop count
                if (selectedSequence.LoopCount == 0)
                {
                    loopForeverCheckbox.Checked = true;
                    iterationsInput.Value = 1; // Default value when loop forever is checked
                }
                else
                {
                    loopForeverCheckbox.Checked = false;
                    iterationsInput.Value = selectedSequence.LoopCount;
                }
                
                startButton.Enabled = true;
                deleteButton.Enabled = true;
                saveButton.Enabled = true;
            }
            else
            {
                startButton.Enabled = false;
                deleteButton.Enabled = false;
                saveButton.Enabled = false;
            }
        }

        private void AddLogMessage(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => AddLogMessage(message)));
                return;
            }
            
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            logTextBox.AppendText($"[{timestamp}] {message}\n");
            logTextBox.ScrollToCaret();
        }

        private void UpdateStatus(string status)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateStatus(status)));
                return;
            }
            
            statusLabel.Text = status;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            isRunning = false;
            
            // Stop recording if active
            if (isRecording && inputRecorder != null)
            {
                inputRecorder.StopRecording();
            }
            
            base.OnFormClosing(e);
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            int selectedIndex = sequencesList.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < macroSequences.Count)
            {
                MacroSequence sequence = macroSequences[selectedIndex];
                
                // Update the sequence's loop count based on UI settings before saving
                if (loopForeverCheckbox.Checked)
                {
                    sequence.LoopCount = 0; // 0 means infinite
                }
                else
                {
                    sequence.LoopCount = (int)iterationsInput.Value;
                }
                
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    string sequencesDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sequences");
                    if (!Directory.Exists(sequencesDir))
                    {
                        Directory.CreateDirectory(sequencesDir);
                    }
                    
                    saveDialog.InitialDirectory = sequencesDir;
                    saveDialog.Filter = "YAML Files (*.yaml)|*.yaml|JSON Files (*.json)|*.json|All Files (*.*)|*.*";
                    saveDialog.DefaultExt = "yaml";
                    saveDialog.FileName = SanitizeFileName(sequence.Name);
                    
                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            // Update sequence name based on file name
                            string fileName = Path.GetFileNameWithoutExtension(saveDialog.FileName);
                            sequence.Name = fileName;
                            
                            // Determine file format based on extension
                            if (saveDialog.FileName.EndsWith(".json"))
                            {
                                // Save as JSON
                                string json = JsonConvert.SerializeObject(sequence, Formatting.Indented);
                                File.WriteAllText(saveDialog.FileName, json);
                            }
                            else if (saveDialog.FileName.EndsWith(".yaml") || saveDialog.FileName.EndsWith(".yml"))
                            {
                                // Save as YAML
                                var serializer = new YamlDotNet.Serialization.SerializerBuilder().Build();
                                string yaml = serializer.Serialize(sequence);
                                File.WriteAllText(saveDialog.FileName, yaml);
                            }
                            else
                            {
                                // Default to YAML if no recognized extension
                                var serializer = new YamlDotNet.Serialization.SerializerBuilder().Build();
                                string yaml = serializer.Serialize(sequence);
                                File.WriteAllText(saveDialog.FileName, yaml);
                            }
                            
                            // Update UI
                            sequencesList.Items[selectedIndex] = sequence.Name;
                            AddLogMessage($"Sequence '{sequence.Name}' saved successfully.");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error saving sequence: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            AddLogMessage($"Error saving sequence: {ex.Message}");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select a sequence to save.", "No Sequence Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            int selectedIndex = sequencesList.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < macroSequences.Count)
            {
                MacroSequence sequence = macroSequences[selectedIndex];
                
                // Confirm deletion
                DialogResult result = MessageBox.Show(
                    $"Are you sure you want to delete the sequence '{sequence.Name}'?",
                    "Confirm Deletion",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    // Check if there's a file to delete
                    string sequencesDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sequences");
                    string fileName = Path.Combine(sequencesDir, SanitizeFileName(sequence.Name) + ".json");
                    
                    // Remove from list
                    macroSequences.RemoveAt(selectedIndex);
                    sequencesList.Items.RemoveAt(selectedIndex);
                    
                    // Delete file if it exists
                    if (File.Exists(fileName))
                    {
                        try
                        {
                            File.Delete(fileName);
                            AddLogMessage($"Sequence file '{Path.GetFileName(fileName)}' deleted.");
                        }
                        catch (Exception ex)
                        {
                            AddLogMessage($"Error deleting sequence file: {ex.Message}");
                        }
                    }
                    
                    AddLogMessage($"Sequence '{sequence.Name}' deleted.");
                    
                    // Select another item if available
                    if (sequencesList.Items.Count > 0)
                    {
                        sequencesList.SelectedIndex = Math.Min(selectedIndex, sequencesList.Items.Count - 1);
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select a sequence to delete.", "No Sequence Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        
        private string SanitizeFileName(string fileName)
        {
            // Remove invalid characters from file name
            char[] invalidChars = Path.GetInvalidFileNameChars();
            return string.Join("_", fileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries)).TrimEnd('.');
        }
    }
}
