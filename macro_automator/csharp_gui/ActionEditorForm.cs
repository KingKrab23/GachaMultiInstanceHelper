using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace MacroAutomatorGUI
{
    public partial class ActionEditorForm : Form
    {
        private MacroAction _action;
        
        // UI Controls
        private ComboBox cboActionType;
        private TableLayoutPanel pnlParameters;
        private Button btnOK;
        private Button btnCancel;
        
        public MacroAction Action => _action;
        
        public ActionEditorForm(MacroAction action = null, bool isNew = true)
        {
            _action = action ?? new MacroAction();
            InitializeComponent();
            PopulateActionTypes();
            CreateParameterControls();
            
            this.Text = isNew ? "Add Macro Action" : "Edit Macro Action";
        }
        
        private void InitializeComponent()
        {
            this.Text = "Add Macro Action";
            this.Size = new Size(500, 400);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(45, 45, 48);
            this.ForeColor = Color.White;
            
            // Action Type ComboBox
            cboActionType = new ComboBox
            {
                Location = new Point(120, 20),
                Width = 250,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            
            Label lblActionType = new Label
            {
                Text = "Action Type:",
                Location = new Point(20, 23),
                AutoSize = true,
                ForeColor = Color.White
            };
            
            // Parameters Panel
            pnlParameters = new TableLayoutPanel
            {
                Location = new Point(20, 70),
                Width = 450,
                Height = 230,
                ColumnCount = 2,
                RowCount = 5,
                AutoScroll = true,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.None
            };
            
            pnlParameters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
            pnlParameters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
            
            // Buttons
            btnOK = new Button
            {
                Text = "OK",
                Width = 90,
                Height = 35,
                Location = new Point(290, 320),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White
            };
            
            btnCancel = new Button
            {
                Text = "Cancel",
                Width = 90,
                Height = 35,
                Location = new Point(390, 320),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White
            };
            
            // Add controls to form
            this.Controls.Add(lblActionType);
            this.Controls.Add(cboActionType);
            this.Controls.Add(pnlParameters);
            this.Controls.Add(btnOK);
            this.Controls.Add(btnCancel);
            
            // Event handlers
            cboActionType.SelectedIndexChanged += ActionTypeChanged;
            btnOK.Click += BtnOK_Click;
            btnCancel.Click += BtnCancel_Click;
        }
        
        private void PopulateActionTypes()
        {
            cboActionType.Items.Clear();
            cboActionType.Items.Add(ActionType.MouseClick);
            cboActionType.Items.Add(ActionType.MouseMove);
            cboActionType.Items.Add(ActionType.KeyPress);
            cboActionType.Items.Add(ActionType.TypeText);
            cboActionType.Items.Add(ActionType.Wait);
            cboActionType.Items.Add(ActionType.AttachToWindow);
            cboActionType.Items.Add(ActionType.Sleep);
            
            // Set the selected type based on existing action
            switch (_action.Type)
            {
                case ActionType.MouseClick:
                    cboActionType.SelectedIndex = 0;
                    break;
                case ActionType.MouseMove:
                    cboActionType.SelectedIndex = 1;
                    break;
                case ActionType.KeyPress:
                    cboActionType.SelectedIndex = 2;
                    break;
                case ActionType.TypeText:
                    cboActionType.SelectedIndex = 3;
                    break;
                case ActionType.Wait:
                    cboActionType.SelectedIndex = 4;
                    break;
                case ActionType.AttachToWindow:
                    cboActionType.SelectedIndex = 5;
                    break;
                case ActionType.Sleep:
                    cboActionType.SelectedIndex = 6;
                    break;
                default:
                    cboActionType.SelectedIndex = 0;
                    break;
            }
        }
        
        private void ActionTypeChanged(object sender, EventArgs e)
        {
            // Update action type
            switch (cboActionType.SelectedIndex)
            {
                case 0:
                    _action.Type = ActionType.MouseClick;
                    break;
                case 1:
                    _action.Type = ActionType.MouseMove;
                    break;
                case 2:
                    _action.Type = ActionType.KeyPress;
                    break;
                case 3:
                    _action.Type = ActionType.TypeText;
                    break;
                case 4:
                    _action.Type = ActionType.Wait;
                    break;
                case 5:
                    _action.Type = ActionType.AttachToWindow;
                    break;
                case 6:
                    _action.Type = ActionType.Sleep;
                    break;
            }
            
            // Reset parameters if needed
            _action.Parameters.Clear();
                
            // Set default parameters based on type
            switch (_action.Type)
            {
                case ActionType.MouseClick:
                    _action.Parameters["x"] = 0;
                    _action.Parameters["y"] = 0;
                    _action.Parameters["button"] = "left";
                    break;
                case ActionType.MouseMove:
                    _action.Parameters["x"] = 0;
                    _action.Parameters["y"] = 0;
                    break;
                case ActionType.KeyPress:
                    _action.Parameters["key"] = "enter";
                    break;
                case ActionType.TypeText:
                    _action.Parameters["text"] = "";
                    break;
                case ActionType.Wait:
                    _action.Parameters["seconds"] = 1.0;
                    break;
                case ActionType.AttachToWindow:
                    _action.Parameters["window_name"] = "";
                    break;
                case ActionType.Sleep:
                    _action.Parameters["milliseconds"] = 1000;
                    break;
            }
            
            // Update parameter panel
            CreateParameterControls();
        }
        
        private void CreateParameterControls()
        {
            pnlParameters.Controls.Clear();
            pnlParameters.RowStyles.Clear();
            
            int rowIndex = 0;
            
            switch (_action.Type)
            {
                case ActionType.MouseClick:
                case ActionType.MouseMove:
                    AddParameterRow("X:", "x", _action.Parameters.ContainsKey("x") ? _action.Parameters["x"].ToString() : "0", rowIndex++);
                    AddParameterRow("Y:", "y", _action.Parameters.ContainsKey("y") ? _action.Parameters["y"].ToString() : "0", rowIndex++);
                    
                    // Add button dropdown
                    if (_action.Type == ActionType.MouseClick)
                    {
                        Label lblButton = new Label { Text = "Button:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
                        ComboBox cboButton = new ComboBox 
                        { 
                            Dock = DockStyle.Fill, 
                            DropDownStyle = ComboBoxStyle.DropDownList,
                            BackColor = Color.FromArgb(60, 60, 63),
                            ForeColor = Color.White,
                            FlatStyle = FlatStyle.Flat
                        };
                        
                        cboButton.Items.Add("left");
                        cboButton.Items.Add("right");
                        cboButton.Items.Add("middle");
                        cboButton.SelectedItem = _action.Parameters.ContainsKey("button") ? _action.Parameters["button"] : "left";
                        cboButton.Tag = "button";
                        cboButton.SelectedIndexChanged += Parameter_Changed;
                        
                        pnlParameters.Controls.Add(lblButton, 0, rowIndex);
                        pnlParameters.Controls.Add(cboButton, 1, rowIndex);
                        pnlParameters.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
                        rowIndex++;
                    }
                    
                    break;
                    
                case ActionType.KeyPress:
                    Label lblKey = new Label { Text = "Key:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
                    ComboBox cboKey = new ComboBox 
                    { 
                        Dock = DockStyle.Fill, 
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        BackColor = Color.FromArgb(60, 60, 63),
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat
                    };
                    
                    // Add common keys
                    string[] commonKeys = { "enter", "tab", "space", "backspace", "esc", "up", "down", "left", "right", "ctrl", "alt", "shift", "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8", "f9", "f10", "f11", "f12", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" };
                    cboKey.Items.AddRange(commonKeys);
                    
                    string keyValue = _action.Parameters.ContainsKey("key") ? _action.Parameters["key"].ToString() : "enter";
                    if (Array.IndexOf(commonKeys, keyValue) >= 0)
                    {
                        cboKey.SelectedItem = keyValue;
                    }
                    else
                    {
                        cboKey.Items.Add(keyValue);
                        cboKey.SelectedItem = keyValue;
                    }
                    
                    cboKey.Tag = "key";
                    cboKey.SelectedIndexChanged += Parameter_Changed;
                    
                    pnlParameters.Controls.Add(lblKey, 0, rowIndex);
                    pnlParameters.Controls.Add(cboKey, 1, rowIndex);
                    pnlParameters.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
                    break;
                    
                case ActionType.TypeText:
                    AddParameterTextBox("Text:", "text", _action.Parameters.ContainsKey("text") ? _action.Parameters["text"].ToString() : "", rowIndex++, 100);
                    break;
                    
                case ActionType.Wait:
                    AddParameterRow("Seconds:", "seconds", _action.Parameters.ContainsKey("seconds") ? _action.Parameters["seconds"].ToString() : "1.0", rowIndex++);
                    break;
                    
                case ActionType.AttachToWindow:
                    AddParameterTextBox("Window Name:", "window_name", _action.Parameters.ContainsKey("window_name") ? _action.Parameters["window_name"].ToString() : "", rowIndex++, 70);
                    break;
                    
                case ActionType.Sleep:
                    AddParameterRow("Milliseconds:", "milliseconds", _action.Parameters.ContainsKey("milliseconds") ? _action.Parameters["milliseconds"].ToString() : "1000", rowIndex++);
                    break;
            }
        }
        
        private void AddParameterRow(string labelText, string paramName, string defaultValue, int rowIndex)
        {
            Label lbl = new Label { Text = labelText, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            TextBox txt = new TextBox 
            { 
                Dock = DockStyle.Fill, 
                Text = defaultValue,
                Tag = paramName,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            txt.TextChanged += Parameter_Changed;
            
            pnlParameters.Controls.Add(lbl, 0, rowIndex);
            pnlParameters.Controls.Add(txt, 1, rowIndex);
            pnlParameters.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
        }
        
        private void AddParameterTextBox(string labelText, string paramName, string defaultValue, int rowIndex, int height)
        {
            Label lbl = new Label { Text = labelText, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            TextBox txt = new TextBox 
            { 
                Dock = DockStyle.Fill, 
                Text = defaultValue,
                Tag = paramName,
                Multiline = height > 40,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            txt.TextChanged += Parameter_Changed;
            
            pnlParameters.Controls.Add(lbl, 0, rowIndex);
            pnlParameters.Controls.Add(txt, 1, rowIndex);
            pnlParameters.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
        }
        
        private void Parameter_Changed(object sender, EventArgs e)
        {
            if (sender is Control control && control.Tag != null)
            {
                string paramName = control.Tag.ToString();
                object value = null;
                
                if (control is TextBox textBox)
                {
                    // Convert to appropriate type based on parameter name
                    if (paramName == "x" || paramName == "y")
                    {
                        if (int.TryParse(textBox.Text, out int intValue))
                        {
                            value = intValue;
                        }
                        else
                        {
                            value = 0;
                        }
                    }
                    else if (paramName == "seconds")
                    {
                        if (double.TryParse(textBox.Text, out double doubleValue))
                        {
                            value = doubleValue;
                        }
                        else
                        {
                            value = 1.0;
                        }
                    }
                    else if (paramName == "milliseconds")
                    {
                        if (int.TryParse(textBox.Text, out int intMilliseconds))
                        {
                            value = intMilliseconds;
                        }
                        else
                        {
                            value = 1000;
                        }
                    }
                    else
                    {
                        value = textBox.Text;
                    }
                }
                else if (control is ComboBox comboBox)
                {
                    value = comboBox.SelectedItem;
                }
                
                if (value != null)
                {
                    _action.Parameters[paramName] = value;
                }
            }
        }
        
        private void BtnOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
