using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace MacroAutomatorGUI
{
    public class SimpleTestForm : Form
    {
        private Button[] testButtons;
        private Label statusLabel;
        private Button startTestButton;
        private Button resetButton;
        private ComboBox testTypeComboBox;
        
        // Test sequence
        private MacroSequence testSequence;
        private bool isRunningTest = false;
        
        public SimpleTestForm()
        {
            InitializeComponent();
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
                Text = "Click on the buttons to test mouse clicks",
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
            
            this.Controls.Add(controlPanel);
            
            // Create test buttons
            CreateTestButtons();
        }
        
        private void CreateTestButtons()
        {
            testButtons = new Button[9];
            
            // Create a 3x3 grid of buttons
            for (int i = 0; i < 9; i++)
            {
                int row = i / 3;
                int col = i % 3;
                
                testButtons[i] = new Button
                {
                    Text = $"Button {i + 1}",
                    Size = new Size(120, 80),
                    Location = new Point(100 + col * 200, 100 + row * 150),
                    BackColor = Color.LightBlue,
                    ForeColor = Color.Black,
                    Font = new Font("Segoe UI", 10),
                    Tag = i + 1
                };
                
                testButtons[i].Click += TestButton_Click;
                this.Controls.Add(testButtons[i]);
            }
        }
        
        private void TestButton_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            button.BackColor = Color.LightGreen;
            button.Text = $"Clicked {button.Tag}!";
            
            statusLabel.Text = $"Button {button.Tag} was clicked";
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
            
            // Reset all buttons
            ResetButtons();
            
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
                    // Click on the center button (index 4)
                    var centerButton = testButtons[4];
                    var centerAction = new MacroAction();
                    centerAction.Type = ActionType.MouseClick;
                    centerAction.Parameters["x"] = centerButton.Location.X + centerButton.Width / 2;
                    centerAction.Parameters["y"] = centerButton.Location.Y + centerButton.Height / 2;
                    centerAction.Parameters["button"] = "left";
                    sequence.AddAction(centerAction);
                    break;
                    
                case "Multiple Click Test":
                    // Click on all buttons in sequence
                    foreach (var button in testButtons)
                    {
                        var clickAction = new MacroAction();
                        clickAction.Type = ActionType.MouseClick;
                        clickAction.Parameters["x"] = button.Location.X + button.Width / 2;
                        clickAction.Parameters["y"] = button.Location.Y + button.Height / 2;
                        clickAction.Parameters["button"] = "left";
                        sequence.AddAction(clickAction);
                        
                        // Add delay between clicks
                        var delayAction = new MacroAction();
                        delayAction.Type = ActionType.Wait;
                        delayAction.Parameters["seconds"] = 0.5;
                        sequence.AddAction(delayAction);
                    }
                    break;
                    
                case "Click Sequence Test":
                    // Click on buttons in a specific pattern (corners and center)
                    int[] pattern = { 0, 2, 8, 6, 4 };
                    
                    foreach (int index in pattern)
                    {
                        var button = testButtons[index];
                        
                        var patternAction = new MacroAction();
                        patternAction.Type = ActionType.MouseClick;
                        patternAction.Parameters["x"] = button.Location.X + button.Width / 2;
                        patternAction.Parameters["y"] = button.Location.Y + button.Height / 2;
                        patternAction.Parameters["button"] = "left";
                        sequence.AddAction(patternAction);
                        
                        // Add delay between clicks
                        var delayAction = new MacroAction();
                        delayAction.Type = ActionType.Wait;
                        delayAction.Parameters["seconds"] = 0.5;
                        sequence.AddAction(delayAction);
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
                            
                            // Convert form-relative coordinates to screen coordinates
                            Point screenPoint = this.PointToScreen(new Point(x, y));
                            
                            // Update status
                            statusLabel.Text = $"Clicking at screen coordinates ({screenPoint.X}, {screenPoint.Y})";
                            
                            // Perform the click using screen coordinates
                            InputSimulator.SimulateMouseClickDirect(screenPoint.X, screenPoint.Y, button);
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
            ResetButtons();
            statusLabel.Text = "Click on the buttons to test mouse clicks";
        }
        
        private void ResetButtons()
        {
            // Reset all buttons
            foreach (var button in testButtons)
            {
                button.BackColor = Color.LightBlue;
                button.Text = $"Button {button.Tag}";
            }
        }
    }
}
