using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace MacroAutomatorGUI
{
    public partial class SequenceEditorForm : Form
    {
        private MacroSequence _sequence;
        private bool _isNew;
        
        // UI Controls
        private TextBox txtName;
        private TextBox txtDescription;
        private NumericUpDown nudDelay;
        private NumericUpDown nudIterations;
        private TrackBar trkDelay;
        private Label lblDelayValue;
        private ListView lstActions;
        private Button btnAddAction;
        private Button btnEditAction;
        private Button btnDeleteAction;
        private Button btnMoveUp;
        private Button btnMoveDown;
        private Button btnSave;
        private Button btnCancel;
        
        public MacroSequence Sequence => _sequence;
        
        public SequenceEditorForm(MacroSequence sequence = null, bool isNew = true)
        {
            _sequence = sequence ?? new MacroSequence();
            _isNew = isNew;
            
            InitializeComponent();
            LoadSequenceData();
        }
        
        private void InitializeComponent()
        {
            this.Text = _isNew ? "Create New Sequence" : "Edit Sequence";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(45, 45, 48);
            this.ForeColor = Color.White;
            
            // Name label and textbox
            Label lblName = new Label
            {
                Text = "Sequence Name:",
                Location = new Point(20, 20),
                AutoSize = true,
                ForeColor = Color.White
            };
            
            txtName = new TextBox
            {
                Location = new Point(150, 17),
                Width = 250,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            
            // Description label and textbox
            Label lblDescription = new Label
            {
                Text = "Sequence Description:",
                Location = new Point(20, 55),
                AutoSize = true,
                ForeColor = Color.White
            };
            
            txtDescription = new TextBox
            {
                Location = new Point(150, 52),
                Width = 250,
                Height = 60,
                Multiline = true,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            
            // Iterations label and numeric up/down
            Label lblIterations = new Label
            {
                Text = "Default Iterations:",
                Location = new Point(450, 20),
                AutoSize = true,
                ForeColor = Color.White
            };
            
            nudIterations = new NumericUpDown
            {
                Location = new Point(580, 17),
                Width = 70,
                Minimum = 1,
                Maximum = 9999,
                Value = 1,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White
            };
            nudIterations.Controls[0].BackColor = Color.FromArgb(60, 60, 63);
            
            // Delay label and numeric up/down
            Label lblDelay = new Label
            {
                Text = "Iteration Delay:",
                Location = new Point(20, 120),
                AutoSize = true,
                ForeColor = Color.White
            };
            
            nudDelay = new NumericUpDown
            {
                Location = new Point(150, 117),
                Width = 100,
                Minimum = 0,
                Maximum = 100,
                DecimalPlaces = 1,
                Value = 0,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White
            };
            nudDelay.Controls[0].BackColor = Color.FromArgb(60, 60, 63);
            
            // Actions list
            Label lblActions = new Label
            {
                Text = "Actions:",
                Location = new Point(20, 155),
                AutoSize = true,
                ForeColor = Color.White
            };
            
            lstActions = new ListView
            {
                Location = new Point(20, 180),
                Width = 760,
                Height = 380,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White
            };
            
            lstActions.Columns.Add("Step", 50);
            lstActions.Columns.Add("Action", 300);
            lstActions.Columns.Add("Parameters", 400);
            
            // Action buttons
            btnAddAction = new Button
            {
                Text = "Add Action",
                Width = 110,
                Height = 35,
                Location = new Point(20, 570),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White
            };
            
            btnEditAction = new Button
            {
                Text = "Edit Action",
                Width = 110,
                Height = 35,
                Location = new Point(140, 570),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White
            };
            
            btnDeleteAction = new Button
            {
                Text = "Delete Action",
                Width = 110,
                Height = 35,
                Location = new Point(260, 570),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White
            };
            
            btnMoveUp = new Button
            {
                Text = "Move Up",
                Width = 110,
                Height = 35,
                Location = new Point(380, 570),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White
            };
            
            btnMoveDown = new Button
            {
                Text = "Move Down",
                Width = 110,
                Height = 35,
                Location = new Point(500, 570),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White
            };
            
            // Bottom buttons
            btnSave = new Button
            {
                Text = "Save",
                Width = 110,
                Height = 35,
                Location = new Point(560, 570),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White
            };
            
            btnCancel = new Button
            {
                Text = "Cancel",
                Width = 110,
                Height = 35,
                Location = new Point(680, 570),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 63),
                ForeColor = Color.White
            };
            
            // Add controls to form
            this.Controls.Add(lblName);
            this.Controls.Add(txtName);
            this.Controls.Add(lblDescription);
            this.Controls.Add(txtDescription);
            this.Controls.Add(lblIterations);
            this.Controls.Add(nudIterations);
            this.Controls.Add(lblDelay);
            this.Controls.Add(nudDelay);
            this.Controls.Add(lblActions);
            this.Controls.Add(lstActions);
            this.Controls.Add(btnAddAction);
            this.Controls.Add(btnEditAction);
            this.Controls.Add(btnDeleteAction);
            this.Controls.Add(btnMoveUp);
            this.Controls.Add(btnMoveDown);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);
            
            // Event handlers
            btnAddAction.Click += BtnAddAction_Click;
            btnEditAction.Click += BtnEditAction_Click;
            btnDeleteAction.Click += BtnDeleteAction_Click;
            btnMoveUp.Click += BtnMoveUp_Click;
            btnMoveDown.Click += BtnMoveDown_Click;
            btnSave.Click += BtnSave_Click;
            btnCancel.Click += BtnCancel_Click;
            lstActions.SelectedIndexChanged += LstActions_SelectedIndexChanged;
            lstActions.DoubleClick += LstActions_DoubleClick;
        }
        
        private void LoadSequenceData()
        {
            txtName.Text = _sequence.Name;
            txtDescription.Text = _sequence.Description;
            nudDelay.Value = (decimal)_sequence.IterationDelay;
            nudIterations.Value = 1; // Default to 1 iteration
            
            // Populate actions list
            RefreshActionsList();
            
            UpdateButtonStates();
        }
        
        private void RefreshActionsList()
        {
            lstActions.Items.Clear();
            
            for (int i = 0; i < _sequence.Actions.Count; i++)
            {
                MacroAction action = _sequence.Actions[i];
                ListViewItem item = new ListViewItem((i + 1).ToString());
                
                item.SubItems.Add(action.Type.ToString());
                
                // Format parameters
                string parameters = "";
                foreach (var param in action.Parameters)
                {
                    parameters += $"{param.Key}={param.Value}, ";
                }
                if (parameters.Length > 2)
                {
                    parameters = parameters.Substring(0, parameters.Length - 2);
                }
                
                item.SubItems.Add(parameters);
                item.Tag = action;
                
                lstActions.Items.Add(item);
            }
        }
        
        private void UpdateButtonStates()
        {
            bool hasSelection = lstActions.SelectedItems.Count > 0;
            int selectedIndex = hasSelection ? lstActions.SelectedItems[0].Index : -1;
            
            btnEditAction.Enabled = hasSelection;
            btnDeleteAction.Enabled = hasSelection;
            btnMoveUp.Enabled = hasSelection && selectedIndex > 0;
            btnMoveDown.Enabled = hasSelection && selectedIndex < lstActions.Items.Count - 1;
        }
        
        private void BtnAddAction_Click(object sender, EventArgs e)
        {
            using (ActionEditorForm form = new ActionEditorForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _sequence.Actions.Add(form.Action);
                    RefreshActionsList();
                }
            }
        }
        
        private void BtnEditAction_Click(object sender, EventArgs e)
        {
            if (lstActions.SelectedItems.Count > 0)
            {
                int index = lstActions.SelectedItems[0].Index;
                MacroAction action = _sequence.Actions[index];
                
                using (ActionEditorForm form = new ActionEditorForm(action, false))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        _sequence.Actions[index] = form.Action;
                        RefreshActionsList();
                        
                        // Re-select the item
                        if (lstActions.Items.Count > index)
                        {
                            lstActions.Items[index].Selected = true;
                        }
                    }
                }
            }
        }
        
        private void BtnDeleteAction_Click(object sender, EventArgs e)
        {
            if (lstActions.SelectedItems.Count > 0)
            {
                int index = lstActions.SelectedItems[0].Index;
                
                if (MessageBox.Show("Are you sure you want to delete this action?", "Confirm Delete", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _sequence.Actions.RemoveAt(index);
                    RefreshActionsList();
                    
                    // Select the next item if available
                    if (lstActions.Items.Count > 0)
                    {
                        int newIndex = Math.Min(index, lstActions.Items.Count - 1);
                        lstActions.Items[newIndex].Selected = true;
                    }
                }
            }
        }
        
        private void BtnMoveUp_Click(object sender, EventArgs e)
        {
            if (lstActions.SelectedItems.Count > 0)
            {
                int index = lstActions.SelectedItems[0].Index;
                
                if (index > 0)
                {
                    MacroAction action = _sequence.Actions[index];
                    _sequence.Actions.RemoveAt(index);
                    _sequence.Actions.Insert(index - 1, action);
                    
                    RefreshActionsList();
                    
                    // Re-select the moved item
                    lstActions.Items[index - 1].Selected = true;
                }
            }
        }
        
        private void BtnMoveDown_Click(object sender, EventArgs e)
        {
            if (lstActions.SelectedItems.Count > 0)
            {
                int index = lstActions.SelectedItems[0].Index;
                
                if (index < lstActions.Items.Count - 1)
                {
                    MacroAction action = _sequence.Actions[index];
                    _sequence.Actions.RemoveAt(index);
                    _sequence.Actions.Insert(index + 1, action);
                    
                    RefreshActionsList();
                    
                    // Re-select the moved item
                    lstActions.Items[index + 1].Selected = true;
                }
            }
        }
        
        private void LstActions_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateButtonStates();
        }
        
        private void LstActions_DoubleClick(object sender, EventArgs e)
        {
            if (lstActions.SelectedItems.Count > 0)
            {
                BtnEditAction_Click(sender, e);
            }
        }
        
        private void BtnSave_Click(object sender, EventArgs e)
        {
            // Validate sequence
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Please enter a name for the sequence.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtName.Focus();
                return;
            }
            
            if (_sequence.Actions.Count == 0)
            {
                if (MessageBox.Show("This sequence has no actions. Do you want to save it anyway?", 
                    "No Actions", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }
            }
            
            // Update sequence properties
            _sequence.Name = txtName.Text;
            _sequence.Description = txtDescription.Text;
            _sequence.IterationDelay = (double)nudDelay.Value;
            // No need to set DefaultIterations as it's no longer in the model
            
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (_isNew || _sequence.Actions.Count == 0 || 
                MessageBox.Show("Are you sure you want to discard your changes?", "Confirm Cancel", 
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }
    }
}
