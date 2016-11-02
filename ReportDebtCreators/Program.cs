using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ReportDebtCreators.enginer;
using ReportDebtCreators.Properties;

namespace ReportDebtCreators
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            RunApp();
        }

        public static void RunApp(bool erFolder=false)
        {

            if (erFolder)
            {
                Dialog(Resources.Fatal_Error_Root_Dir, Resources.atal_Error_Root_Dir_Title);
                return;
            }

            var dialog = new FolderBrowserDialog
            {
                ShowNewFolderButton = false
            };



            if (dialog.ShowDialog() == DialogResult.Cancel)
            {
                Dialog(Resources.Fatal_Error_Root_Dir, Resources.atal_Error_Root_Dir_Title);
            }
            else
            {

                var getData = new ExelCreator(dialog.SelectedPath);

                var temp = getData.ListTemplate();
                var pack = getData.ListPackage();

                if (temp == null || pack == null)
                {
                    RunApp(true);
                }
                else
                {
                    Application.Run(new MainCreatorsForm(temp, pack));
                }

            }
            
        }

        private static void Dialog(string title, string text)
        {
            var reflex = MessageBox.Show(title, text, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            if (reflex == DialogResult.Cancel) ExitApp();
            if (reflex == DialogResult.Retry)
            {
                RunApp();
            }
        }

        public static void ExitApp()
        {
            Application.Exit();
        }
    }
}
