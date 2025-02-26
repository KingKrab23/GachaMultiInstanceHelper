using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

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
                Text = "Stop (Esc)",
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
                RowCount = 2,
                ColumnCount = 1,
                BackColor = Color.FromArgb(35, 35, 35)
            };
            
            sequencesLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            sequencesLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            sequencesPanel.Controls.Add(sequencesLayout);

            Label sequencesLabel = new Label
            {
                Text = "Available Sequences",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                AutoSize = true,
                Dock = DockStyle.Fill
            };
            sequencesLayout.Controls.Add(sequencesLabel, 0, 0);

            sequencesList = new ListBox
            {
                BackColor = Color.FromArgb(50, 50, 50),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F)
            };
            sequencesLayout.Controls.Add(sequencesList, 0, 1);

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
                else if (e.KeyCode == Keys.Escape)
                {
                    StopButton_Click(sender, e);
                }
            };
        }

        private void LoadSequences()
        {
            AddLogMessage("Loading sequences...");
            
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
            int iterations = loopForeverCheckbox.Checked ? 0 : (int)iterationsInput.Value;

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
    }
}
