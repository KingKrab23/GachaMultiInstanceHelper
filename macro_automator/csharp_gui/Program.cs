using System;
using System.Windows.Forms;

namespace MacroAutomatorGUI
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Check for command line arguments
            if (args.Length > 0)
            {
                switch (args[0].ToLower())
                {
                    case "--test":
                    case "-t":
                        // Launch the test form
                        Application.Run(new SimpleTestForm());
                        return;
                        
                    case "--help":
                    case "-h":
                        ShowHelp();
                        return;
                }
            }

            // Default: launch the main application
            Application.Run(new MainFormSimplified());
        }
        
        private static void ShowHelp()
        {
            string helpText = 
                "Macro Automator Command Line Options:\n\n" +
                "--test, -t    Launch the mouse click test form\n" +
                "--help, -h    Show this help message\n";
                
            MessageBox.Show(helpText, "Macro Automator Help", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
