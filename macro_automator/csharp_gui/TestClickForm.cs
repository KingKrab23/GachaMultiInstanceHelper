using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace MacroAutomatorGUI
{
    public class TestClickForm : Form
    {
        private List<ClickTarget> clickTargets = new List<ClickTarget>();
        private Label statusLabel;
        private Button startTestButton;
        private Button resetButton;
        private ComboBox testTypeComboBox;
        private CheckBox showCoordinatesCheckBox;
        private System.Windows.Forms.Timer coordinateTimer;
        private Label coordinateLabel;
        private System.ComponentModel.IContainer components = null;
        
        // Test sequence
        private MacroSequence testSequence;
        private bool isRunningTest = false;
        
        public TestClickForm()
        {
            InitializeComponent();
            CreateClickTargets();
            SetupTestSequences();
        }
        
        private void InitializeComponent()
        {
            // Form settings
            this.Text = "Mouse Click Test";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(240, 240, 240);
            this.ForeColor = Color.Black;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            
            // Status label
            statusLabel = new Label
            {
                Text = "Click on the targets to test mouse clicks",
                Dock = DockStyle.Bottom,
                Height = 30,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            this.Controls.Add(statusLabel);
            
            // Control panel
            Panel controlPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                Padding = new Padding(10)
            };
            
            // Test type combo box
            testTypeComboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 200,
                Location = new Point(10, 15),
                Font = new Font("Segoe UI", 9)
            };
            testTypeComboBox.Items.AddRange(new object[] { 
                "Basic Click Test", 
                "Multiple Click Test", 
                "Different Button Test",
                "Click and Drag Test",
                "Click Sequence Test"
            });
            testTypeComboBox.SelectedIndex = 0;
            controlPanel.Controls.Add(testTypeComboBox);
            
            // Start test button
            startTestButton = new Button
            {
                Text = "Start Test",
                Location = new Point(220, 15),
                Width = 100,
                Font = new Font("Segoe UI", 9)
            };
            startTestButton.Click += StartTest_Click;
            controlPanel.Controls.Add(startTestButton);
            
            // Reset button
            resetButton = new Button
            {
                Text = "Reset",
                Location = new Point(330, 15),
                Width = 100,
                Font = new Font("Segoe UI", 9)
            };
            resetButton.Click += ResetButton_Click;
            controlPanel.Controls.Add(resetButton);
            
            // Show coordinates checkbox
            showCoordinatesCheckBox = new CheckBox
            {
                Text = "Show Mouse Coordinates",
                Location = new Point(440, 15),
                Width = 180,
                Font = new Font("Segoe UI", 9)
            };
            showCoordinatesCheckBox.CheckedChanged += ShowCoordinatesCheckBox_CheckedChanged;
            controlPanel.Controls.Add(showCoordinatesCheckBox);
            
            // Coordinate label
            coordinateLabel = new Label
            {
                Text = "X: 0, Y: 0",
                Location = new Point(630, 15),
                Width = 150,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9),
                Visible = false
            };
            controlPanel.Controls.Add(coordinateLabel);
            
            // Coordinate timer
            coordinateTimer = new System.Windows.Forms.Timer()
            {
                Interval = 50,
                Enabled = false
            };
            coordinateTimer.Tick += CoordinateTimer_Tick;
            
            this.Controls.Add(controlPanel);
            
            // Set up mouse events for the form
            this.MouseClick += TestClickForm_MouseClick;
        }
        
        private void CreateClickTargets()
        {
            // Create click targets at different positions
            clickTargets.Add(new ClickTarget(100, 150, "Target 1"));
            clickTargets.Add(new ClickTarget(300, 150, "Target 2"));
            clickTargets.Add(new ClickTarget(500, 150, "Target 3"));
            clickTargets.Add(new ClickTarget(100, 300, "Target 4"));
            clickTargets.Add(new ClickTarget(300, 300, "Target 5"));
            clickTargets.Add(new ClickTarget(500, 300, "Target 6"));
            clickTargets.Add(new ClickTarget(100, 450, "Target 7"));
            clickTargets.Add(new ClickTarget(300, 450, "Target 8"));
            clickTargets.Add(new ClickTarget(500, 450, "Target 9"));
            
            // Add targets to form
            foreach (var target in clickTargets)
            {
                this.Controls.Add(target);
            }
        }
        
        private void SetupTestSequences()
        {
            // Create test sequences for different test types
            testSequence = new MacroSequence("Test Sequence");
        }
        
        private void StartTest_Click(object sender, EventArgs e)
        {
            if (isRunningTest)
            {
                return;
            }
            
            isRunningTest = true;
            statusLabel.Text = "Running test...";
            
            // Reset all targets
            foreach (var target in clickTargets)
            {
                target.Reset();
            }
            
            // Create and run test based on selected type
            string testType = testTypeComboBox.SelectedItem.ToString();
            
            Task.Run(() => RunTest(testType));
        }
        
        private async Task RunTest(string testType)
        {
            // Wait a moment before starting the test
            await Task.Delay(1000);
            
            MacroSequence sequence = new MacroSequence($"{testType}");
            
            switch (testType)
            {
                case "Basic Click Test":
                    // Click on the center target
                    var clickAction = new MacroAction();
                    clickAction.Type = ActionType.MouseClick;
                    clickAction.Parameters["x"] = 300;
                    clickAction.Parameters["y"] = 300;
                    clickAction.Parameters["button"] = "left";
                    sequence.AddAction(clickAction);
                    break;
                    
                case "Multiple Click Test":
                    // Click on multiple targets in sequence
                    foreach (var target in clickTargets)
                    {
                        var multiClickAction = new MacroAction();
                        multiClickAction.Type = ActionType.MouseClick;
                        multiClickAction.Parameters["x"] = target.CenterX;
                        multiClickAction.Parameters["y"] = target.CenterY;
                        multiClickAction.Parameters["button"] = "left";
                        sequence.AddAction(multiClickAction);
                        
                        // Add delay between clicks
                        var delayAction = new MacroAction();
                        delayAction.Type = ActionType.Wait;
                        delayAction.Parameters["seconds"] = 0.5;
                        sequence.AddAction(delayAction);
                    }
                    break;
                    
                case "Different Button Test":
                    // Test different mouse buttons
                    // Left click
                    var leftClickAction = new MacroAction();
                    leftClickAction.Type = ActionType.MouseClick;
                    leftClickAction.Parameters["x"] = 100;
                    leftClickAction.Parameters["y"] = 150;
                    leftClickAction.Parameters["button"] = "left";
                    sequence.AddAction(leftClickAction);
                    
                    var leftDelayAction = new MacroAction();
                    leftDelayAction.Type = ActionType.Wait;
                    leftDelayAction.Parameters["seconds"] = 0.5;
                    sequence.AddAction(leftDelayAction);
                    
                    // Right click
                    var rightClickAction = new MacroAction();
                    rightClickAction.Type = ActionType.MouseClick;
                    rightClickAction.Parameters["x"] = 300;
                    rightClickAction.Parameters["y"] = 150;
                    rightClickAction.Parameters["button"] = "right";
                    sequence.AddAction(rightClickAction);
                    
                    var rightDelayAction = new MacroAction();
                    rightDelayAction.Type = ActionType.Wait;
                    rightDelayAction.Parameters["seconds"] = 0.5;
                    sequence.AddAction(rightDelayAction);
                    
                    // Middle click
                    var middleClickAction = new MacroAction();
                    middleClickAction.Type = ActionType.MouseClick;
                    middleClickAction.Parameters["x"] = 500;
                    middleClickAction.Parameters["y"] = 150;
                    middleClickAction.Parameters["button"] = "middle";
                    sequence.AddAction(middleClickAction);
                    break;
                    
                case "Click and Drag Test":
                    // Move to first position
                    var moveFirstAction = new MacroAction();
                    moveFirstAction.Type = ActionType.MouseMove;
                    moveFirstAction.Parameters["x"] = 100;
                    moveFirstAction.Parameters["y"] = 300;
                    sequence.AddAction(moveFirstAction);
                    
                    // Mouse down
                    var mouseDownAction = new MacroAction();
                    mouseDownAction.Type = ActionType.MouseClick;
                    mouseDownAction.Parameters["x"] = 100;
                    mouseDownAction.Parameters["y"] = 300;
                    mouseDownAction.Parameters["button"] = "left";
                    sequence.AddAction(mouseDownAction);
                    
                    // Move to second position
                    var moveSecondAction = new MacroAction();
                    moveSecondAction.Type = ActionType.MouseMove;
                    moveSecondAction.Parameters["x"] = 500;
                    moveSecondAction.Parameters["y"] = 300;
                    sequence.AddAction(moveSecondAction);
                    
                    // Mouse up
                    var mouseUpAction = new MacroAction();
                    mouseUpAction.Type = ActionType.MouseClick;
                    mouseUpAction.Parameters["x"] = 500;
                    mouseUpAction.Parameters["y"] = 300;
                    mouseUpAction.Parameters["button"] = "left";
                    sequence.AddAction(mouseUpAction);
                    break;
                    
                case "Click Sequence Test":
                    // Click on targets in a specific pattern
                    int[] pattern = { 0, 2, 8, 6, 4 }; // Corners and center
                    
                    foreach (int index in pattern)
                    {
                        if (index < clickTargets.Count)
                        {
                            var target = clickTargets[index];
                            
                            var sequenceClickAction = new MacroAction();
                            sequenceClickAction.Type = ActionType.MouseClick;
                            sequenceClickAction.Parameters["x"] = target.CenterX;
                            sequenceClickAction.Parameters["y"] = target.CenterY;
                            sequenceClickAction.Parameters["button"] = "left";
                            sequence.AddAction(sequenceClickAction);
                            
                            // Add delay between clicks
                            var sequenceDelayAction = new MacroAction();
                            sequenceDelayAction.Type = ActionType.Wait;
                            sequenceDelayAction.Parameters["seconds"] = 0.5;
                            sequence.AddAction(sequenceDelayAction);
                        }
                    }
                    break;
            }
            
            // Execute the test sequence
            await ExecuteTestSequence(sequence);
            
            // Update UI when test is complete
            this.Invoke(new Action(() => {
                isRunningTest = false;
                statusLabel.Text = "Test completed";
            }));
        }
        
        private async Task ExecuteTestSequence(MacroSequence sequence)
        {
            foreach (var action in sequence.Actions)
            {
                await this.Invoke(async () => {
                    switch (action.Type)
                    {
                        case ActionType.MouseClick:
                            int x = Convert.ToInt32(action.Parameters["x"]);
                            int y = Convert.ToInt32(action.Parameters["y"]);
                            string button = action.Parameters["button"].ToString();
                            
                            // Update status
                            statusLabel.Text = $"Clicking at ({x}, {y}) with {button} button";
                            
                            // Perform the click
                            InputSimulator.SimulateMouseClickDirect(x, y, button);
                            break;
                            
                        case ActionType.MouseMove:
                            int moveX = Convert.ToInt32(action.Parameters["x"]);
                            int moveY = Convert.ToInt32(action.Parameters["y"]);
                            
                            // Update status
                            statusLabel.Text = $"Moving mouse to ({moveX}, {moveY})";
                            
                            // Perform the move
                            InputSimulator.SimulateMouseMove(moveX, moveY);
                            break;
                            
                        case ActionType.Wait:
                            double seconds = Convert.ToDouble(action.Parameters["seconds"]);
                            
                            // Update status
                            statusLabel.Text = $"Waiting for {seconds} seconds";
                            
                            // Wait
                            await Task.Delay((int)(seconds * 1000));
                            break;
                    }
                });
            }
        }
        
        private void ResetButton_Click(object sender, EventArgs e)
        {
            // Reset all targets
            foreach (var target in clickTargets)
            {
                target.Reset();
            }
            
            statusLabel.Text = "Click on the targets to test mouse clicks";
        }
        
        private void ShowCoordinatesCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            coordinateLabel.Visible = showCoordinatesCheckBox.Checked;
            coordinateTimer.Enabled = showCoordinatesCheckBox.Checked;
        }
        
        private void CoordinateTimer_Tick(object sender, EventArgs e)
        {
            // Update coordinate display
            Point mousePos = this.PointToClient(Cursor.Position);
            coordinateLabel.Text = $"X: {mousePos.X}, Y: {mousePos.Y}";
        }
        
        private void TestClickForm_MouseClick(object sender, MouseEventArgs e)
        {
            // Check if any target was clicked
            bool targetClicked = false;
            
            foreach (var target in clickTargets)
            {
                if (target.Contains(e.Location))
                {
                    target.Click();
                    statusLabel.Text = $"Clicked on {target.Name}";
                    targetClicked = true;
                    break;
                }
            }
            
            if (!targetClicked)
            {
                statusLabel.Text = $"Clicked at ({e.X}, {e.Y}) - No target hit";
            }
        }
    }
    
    // Click target control
    public class ClickTarget : Control
    {
        private bool isClicked = false;
        private string targetName;
        
        public int CenterX => this.Left + this.Width / 2;
        public int CenterY => this.Top + this.Height / 2;
        public new string Name => targetName;
        
        public ClickTarget(int x, int y, string name)
        {
            this.targetName = name;
            this.Size = new Size(80, 80);
            this.Location = new Point(x - this.Width / 2, y - this.Height / 2);
            this.BackColor = Color.LightBlue;
            this.ForeColor = Color.Black;
            this.Text = name;
            this.Font = new Font("Segoe UI", 9);
            
            // Set up mouse events
            this.MouseClick += ClickTarget_MouseClick;
        }
        
        private void ClickTarget_MouseClick(object sender, MouseEventArgs e)
        {
            Click();
        }
        
        public new void Click()
        {
            isClicked = true;
            this.BackColor = Color.LightGreen;
            this.Text = $"{targetName}\nClicked!";
        }
        
        public void Reset()
        {
            isClicked = false;
            this.BackColor = Color.LightBlue;
            this.Text = targetName;
        }
        
        public bool Contains(Point p)
        {
            return p.X >= this.Left && p.X <= this.Right &&
                   p.Y >= this.Top && p.Y <= this.Bottom;
        }
        
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            
            // Draw border
            using (Pen pen = new Pen(Color.DarkBlue, 2))
            {
                e.Graphics.DrawRectangle(pen, 0, 0, this.Width - 1, this.Height - 1);
            }
            
            // Draw target lines
            if (!isClicked)
            {
                using (Pen pen = new Pen(Color.DarkBlue, 1))
                {
                    int center = this.Width / 2;
                    e.Graphics.DrawLine(pen, center, 10, center, this.Height - 10);
                    e.Graphics.DrawLine(pen, 10, center, this.Width - 10, center);
                }
            }
        }
    }
}
