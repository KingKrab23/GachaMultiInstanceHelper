using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MacroAutomatorGUI
{
    public partial class MainForm : Form
    {
        private Process pythonProcess;
        private CancellationTokenSource cancellationTokenSource;
        private List<MacroSequence> sequences = new List<MacroSequence>();
        private bool isRunning = false;
        private string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sequences.yaml");
        private string pythonScriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "macro_automator.py");

        public MainForm(string configPath = null, bool autoStart = false, string startSequence = null, int iterations = 1)
        {
            if (configPath != null)
            {
                this.configPath = configPath;
            }

            InitializeComponent();
            LoadSequences();
            SetupHotkeys();

            // Auto-start sequence if requested
            if (autoStart && startSequence != null)
            {
                var sequencesList = this.Controls.Find("sequencesList", true)[0] as ListBox;
                var iterationsInput = this.Controls.Find("iterationsInput", true)[0] as NumericUpDown;
                var loopForeverCheckbox = this.Controls.Find("loopForeverCheckbox", true)[0] as CheckBox;

                // Find the requested sequence
                for (int i = 0; i < sequencesList.Items.Count; i++)
                {
                    if (sequencesList.Items[i].ToString() == startSequence)
                    {
                        sequencesList.SelectedIndex = i;
                        iterationsInput.Value = iterations;
                        loopForeverCheckbox.Checked = iterations <= 0;
                        
                        // Start the sequence after a short delay to allow the form to load
                        Task.Delay(500).ContinueWith(_ => 
                        {
                            this.Invoke(new Action(() => StartButton_Click(null, EventArgs.Empty)));
                        });
                        
                        break;
                    }
                }
            }
        }

        private void InitializeComponent()
        {
            // Form settings
            this.Text = "Macro Automator";
            this.Size = new Size(700, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.Icon = SystemIcons.Application;
            this.BackColor = Color.FromArgb(30, 30, 30);
            this.ForeColor = Color.White;

            // Main layout
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.White,
                RowCount = 3,
                ColumnCount = 1
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 120));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 150));
            this.Controls.Add(mainLayout);

            // Top panel (controls)
            Panel controlPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(40, 40, 40),
                Padding = new Padding(10)
            };
            mainLayout.Controls.Add(controlPanel, 0, 0);

            // Control buttons
            Button startButton = new Button
            {
                Text = "Start Sequence (F5)",
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(150, 40),
                Location = new Point(20, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            startButton.Click += StartButton_Click;
            controlPanel.Controls.Add(startButton);

            Button stopButton = new Button
            {
                Text = "Stop (Esc)",
                BackColor = Color.FromArgb(200, 0, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(150, 40),
                Location = new Point(180, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            stopButton.Click += StopButton_Click;
            controlPanel.Controls.Add(stopButton);

            Button newSequenceButton = new Button
            {
                Text = "New Sequence",
                BackColor = Color.FromArgb(0, 150, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(150, 40),
                Location = new Point(340, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            newSequenceButton.Click += NewSequenceButton_Click;
            controlPanel.Controls.Add(newSequenceButton);

            Label iterationsLabel = new Label
            {
                Text = "Iterations:",
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(20, 75)
            };
            controlPanel.Controls.Add(iterationsLabel);

            NumericUpDown iterationsInput = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 9999,
                Value = 1,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                Size = new Size(80, 25),
                Location = new Point(90, 73),
                Name = "iterationsInput"
            };
            controlPanel.Controls.Add(iterationsInput);

            CheckBox loopForeverCheckbox = new CheckBox
            {
                Text = "Loop Forever",
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(180, 75),
                Name = "loopForeverCheckbox"
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

            Label sequencesLabel = new Label
            {
                Text = "Available Sequences",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(10, 10)
            };
            sequencesPanel.Controls.Add(sequencesLabel);

            ListBox sequencesList = new ListBox
            {
                BackColor = Color.FromArgb(50, 50, 50),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill,
                Name = "sequencesList",
                Font = new Font("Consolas", 10F),
                Margin = new Padding(10, 40, 10, 10)
            };
            sequencesList.SelectedIndexChanged += SequencesList_SelectedIndexChanged;
            sequencesPanel.Controls.Add(sequencesList);
            sequencesList.BringToFront();

            // Bottom panel (log)
            Panel logPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(25, 25, 25),
                Padding = new Padding(10)
            };
            mainLayout.Controls.Add(logPanel, 0, 2);

            Label logLabel = new Label
            {
                Text = "Log",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(10, 10)
            };
            logPanel.Controls.Add(logLabel);

            RichTextBox logTextBox = new RichTextBox
            {
                BackColor = Color.FromArgb(45, 45, 45),
                ForeColor = Color.LightGray,
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Name = "logTextBox",
                Font = new Font("Consolas", 9F),
                Margin = new Padding(10, 35, 10, 10)
            };
            logPanel.Controls.Add(logTextBox);
            logTextBox.BringToFront();
            
            // Add status strip
            StatusStrip statusStrip = new StatusStrip
            {
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.White,
                SizingGrip = false
            };
            
            ToolStripStatusLabel statusLabel = new ToolStripStatusLabel
            {
                Text = "Ready",
                ForeColor = Color.White
            };
            
            statusStrip.Items.Add(statusLabel);
            this.Controls.Add(statusStrip);

            // Add to form
            this.Controls.Add(mainLayout);
        }

        private void SetupHotkeys()
        {
            // Set up global hotkeys
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
            // In a real application, this would load sequences from the config file
            AddLogMessage("Loading sequences...");
            
            if (File.Exists(configPath))
            {
                // Here we would parse the YAML file
                AddLogMessage($"Loaded sequences from {configPath}");
            }
            else
            {
                // Add some example sequences
                sequences.Add(new MacroSequence
                {
                    Name = "Click Center",
                    Description = "Clicks at the center of the active window",
                    Actions = new List<MacroAction>
                    {
                        new MacroAction { Type = ActionType.MouseClick, Parameters = new Dictionary<string, object>() }
                    }
                });
                
                sequences.Add(new MacroSequence
                {
                    Name = "Type Hello World",
                    Description = "Types 'Hello World' in the active window",
                    Actions = new List<MacroAction>
                    {
                        new MacroAction
                        {
                            Type = ActionType.TypeText,
                            Parameters = new Dictionary<string, object> { { "text", "Hello World" } }
                        }
                    }
                });
                
                sequences.Add(new MacroSequence
                {
                    Name = "Press Enter 5 Times",
                    Description = "Presses Enter key 5 times with delays",
                    Actions = new List<MacroAction>
                    {
                        new MacroAction { Type = ActionType.KeyPress, Parameters = new Dictionary<string, object> { { "key", "enter" } } },
                        new MacroAction { Type = ActionType.Sleep, Parameters = new Dictionary<string, object> { { "milliseconds", 500 } } },
                        new MacroAction { Type = ActionType.KeyPress, Parameters = new Dictionary<string, object> { { "key", "enter" } } },
                        new MacroAction { Type = ActionType.Sleep, Parameters = new Dictionary<string, object> { { "milliseconds", 500 } } },
                        new MacroAction { Type = ActionType.KeyPress, Parameters = new Dictionary<string, object> { { "key", "enter" } } },
                        new MacroAction { Type = ActionType.Sleep, Parameters = new Dictionary<string, object> { { "milliseconds", 500 } } },
                        new MacroAction { Type = ActionType.KeyPress, Parameters = new Dictionary<string, object> { { "key", "enter" } } },
                        new MacroAction { Type = ActionType.Sleep, Parameters = new Dictionary<string, object> { { "milliseconds", 500 } } },
                        new MacroAction { Type = ActionType.KeyPress, Parameters = new Dictionary<string, object> { { "key", "enter" } } }
                    }
                });
                
                AddLogMessage("Created example sequences");
            }

            // Update the list
            ListBox sequencesList = this.Controls.Find("sequencesList", true)[0] as ListBox;
            sequencesList.Items.Clear();
            foreach (var sequence in sequences)
            {
                sequencesList.Items.Add($"{sequence.Name} - {sequence.Description}");
            }
            
            if (sequencesList.Items.Count > 0)
            {
                sequencesList.SelectedIndex = 0;
            }
            
            AddLogMessage($"Loaded {sequences.Count} sequences");
            UpdateStatus("Ready");
        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            if (isRunning)
            {
                AddLogMessage("A sequence is already running");
                return;
            }

            ListBox sequencesList = this.Controls.Find("sequencesList", true)[0] as ListBox;
            if (sequencesList.SelectedIndex < 0)
            {
                AddLogMessage("No sequence selected");
                return;
            }

            NumericUpDown iterationsInput = this.Controls.Find("iterationsInput", true)[0] as NumericUpDown;
            CheckBox loopForeverCheckbox = this.Controls.Find("loopForeverCheckbox", true)[0] as CheckBox;
            
            int iterations = loopForeverCheckbox.Checked ? -1 : (int)iterationsInput.Value;
            string sequenceName = sequences[sequencesList.SelectedIndex].Name;
            
            ExecuteSequence(sequenceName, iterations);
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            if (!isRunning)
            {
                AddLogMessage("No sequence is currently running");
                return;
            }

            StopExecution();
        }

        private void NewSequenceButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This feature is not implemented in this demo.\n\nIn a full application, this would open a sequence editor.", 
                "Feature Not Implemented", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SequencesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListBox sequencesList = sender as ListBox;
            if (sequencesList.SelectedIndex >= 0)
            {
                var sequence = sequences[sequencesList.SelectedIndex];
                AddLogMessage($"Selected sequence: {sequence.Name}");
                AddLogMessage($"Description: {sequence.Description}");
                AddLogMessage($"Actions: {sequence.Actions.Count}");
            }
        }

        private async void ExecuteSequence(string sequenceName, int iterations)
        {
            try
            {
                isRunning = true;
                UpdateStatus($"Running sequence: {sequenceName}");
                AddLogMessage($"Starting sequence: {sequenceName} with {(iterations < 0 ? "infinite" : iterations.ToString())} iterations");
                
                // In a real application, this would call the Python script
                // Here we just simulate the execution
                await Task.Run(() => 
                {
                    // Find the selected sequence
                    var sequence = sequences.Find(s => s.Name == sequenceName);
                    if (sequence == null) return;
                    
                    // Execute the actions
                    int iterCount = 0;
                    while ((iterations < 0 || iterCount < iterations) && isRunning)
                    {
                        iterCount++;
                        this.Invoke(new Action(() => 
                        {
                            AddLogMessage($"Iteration {iterCount}" + (iterations < 0 ? "" : $"/{iterations}"));
                        }));
                        
                        foreach (var action in sequence.Actions)
                        {
                            if (!isRunning) break;
                            
                            this.Invoke(new Action(() => 
                            {
                                AddLogMessage($"  Executing: {action.Type}");
                            }));
                            
                            // Simulate action execution time
                            if (action.Type == ActionType.Sleep && action.Parameters.ContainsKey("milliseconds"))
                            {
                                double duration = Convert.ToDouble(action.Parameters["milliseconds"]);
                                System.Threading.Thread.Sleep((int)(duration));
                            }
                            else
                            {
                                // Other actions would take some time
                                System.Threading.Thread.Sleep(100);
                            }
                        }
                        
                        // If running more iterations, pause between them
                        if (isRunning && (iterations < 0 || iterCount < iterations))
                        {
                            System.Threading.Thread.Sleep(500);
                        }
                    }
                });
                
                // Execution completed or was stopped
                if (isRunning)
                {
                    AddLogMessage("Sequence execution completed");
                }
                else
                {
                    AddLogMessage("Sequence execution was stopped");
                }
            }
            catch (Exception ex)
            {
                AddLogMessage($"Error executing sequence: {ex.Message}");
            }
            finally
            {
                isRunning = false;
                UpdateStatus("Ready");
            }
        }

        private void StopExecution()
        {
            if (!isRunning) return;
            
            AddLogMessage("Stopping sequence execution...");
            isRunning = false;
            
            if (pythonProcess != null && !pythonProcess.HasExited)
            {
                try
                {
                    pythonProcess.Kill();
                    pythonProcess = null;
                }
                catch
                {
                    // Ignore errors on process kill
                }
            }
        }

        private void AddLogMessage(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => AddLogMessage(message)));
                return;
            }
            
            RichTextBox logTextBox = this.Controls.Find("logTextBox", true)[0] as RichTextBox;
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
            
            StatusStrip statusStrip = this.Controls[this.Controls.Count - 1] as StatusStrip;
            ToolStripStatusLabel statusLabel = statusStrip.Items[0] as ToolStripStatusLabel;
            statusLabel.Text = status;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            StopExecution();
            base.OnFormClosing(e);
        }
    }
}
