using System;
using System.Configuration;
using System.IO;
using System.Windows.Forms;
using ReportDebtCreators.enginer;
using ReportDebtCreators.Properties;


namespace ReportDebtCreators
{
    static class Program
    {

        private static string _dirTemp;
        private static string _dirEng;
        private static string _dirResRep;
        private static string _dirCompil;
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var defaultPach = ConfigurationManager.AppSettings["rootPachExel"];

            if (string.IsNullOrEmpty(defaultPach)) RunApp();
            else DetectRunApp(defaultPach,true);
        }

        private static void RunApp(bool erFolder=false)
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
                DetectRunApp(dialog.SelectedPath);
            }
            
        }

        private static void Dialog(string title, string text)
        {
            var reflex = MessageBox.Show(title, text, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
            if (reflex == DialogResult.Cancel) Application.Exit();
            if (reflex == DialogResult.Retry)
            {
                RunApp();
            }
        }

        private static void DetectRunApp(string rootPatch, bool isConfig=false)
        {
            var getData = new ExelCreator();

            _dirTemp = $@"{rootPatch}\{ConfigurationManager.AppSettings["dirTemp"]}\";
            _dirEng = $@"{rootPatch}\{ConfigurationManager.AppSettings["dirEng"]}\";
            _dirResRep = $@"{rootPatch}\{ConfigurationManager.AppSettings["dirResRep"]}\";
            _dirCompil = $@"{rootPatch}\{ConfigurationManager.AppSettings["dirCompil"]}\";

            var dirExist = (
                Directory.Exists(_dirTemp) &&
                Directory.Exists(_dirEng) &&
                Directory.Exists(_dirResRep) &&
                Directory.Exists(_dirCompil)
                );

            if(!dirExist) { RunApp(true); return;}

            var temp = getData.ListTemplate(_dirTemp);
            var pack = getData.ListPackage(_dirEng);

            if (temp == null || pack == null)
            {
                RunApp(true);
            }
            else
            {
                if(!isConfig) UpdateSetting("rootPachExel", rootPatch);
                
                Application.Run(new MainCreatorsForm(temp, pack));
            }
        }

        private static void UpdateSetting(string key, string value)
        {
            var configuration = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
            configuration.AppSettings.Settings[key].Value = value;
            configuration.Save();

            //ConfigurationManager.RefreshSection("appSettings");
        }
    }
}
