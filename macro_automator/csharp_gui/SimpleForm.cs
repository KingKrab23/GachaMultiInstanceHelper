using System;
using System.Windows.Forms;

namespace MacroAutomatorGUI
{
    public class SimpleForm : Form
    {
        public SimpleForm()
        {
            this.Text = "Simple Macro Automator";
            this.Size = new System.Drawing.Size(400, 300);
            
            Button button = new Button
            {
                Text = "Click Me",
                Location = new System.Drawing.Point(150, 100),
                Size = new System.Drawing.Size(100, 30)
            };
            
            button.Click += (sender, e) => MessageBox.Show("Button clicked!");
            
            this.Controls.Add(button);
        }
    }
}
