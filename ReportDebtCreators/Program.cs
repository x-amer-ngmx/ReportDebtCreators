using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ReportDebtCreators.enginer;
using ReportDebtCreators.Properties;


namespace ReportDebtCreators
{
    static class Program
    {

        public static string DirTemp { private set; get; }
        public static string DirEng { private set; get; }
        public static string DirResRep { private set; get; }
        public static string DirCompil { private set; get; }

        public static string Pws { private set; get; }
        public static string RootPatch { private set; get; }

        public static int[] cellRange { private set; get; }

        public static string TempJsonDB { private set; get; }

        public static string brnCell { private set; get; }


        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            RootPatch = ConfigurationManager.AppSettings["rootPachExel"];
            Pws = ConfigurationManager.AppSettings["passwordSheet"];
            var _cellR = ConfigurationManager.AppSettings["rangeCell"];
            var _bcell = ConfigurationManager.AppSettings["brangeCell"];

            TempJsonDB = $"{RootPatch}\\Temp_Data.json";

            brnCell = string.IsNullOrEmpty(_bcell) ? "W" : _bcell;

            _cellR = string.IsNullOrEmpty(_cellR) ? "2,3,5,7,8,9,10,14,22,27,28,29,30" : _cellR;

            try
            {
                cellRange = (from x in _cellR.Split(',') select int.Parse(x)).ToArray();
            }
            catch (Exception)
            {
                cellRange = new[] {2, 3, 5, 7, 8, 9, 10, 14, 22, 27, 28, 29, 30};
            }

            
            if (string.IsNullOrEmpty(RootPatch)) RunApp();
            else DetectRunApp(true);
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
                RootPatch = dialog.SelectedPath;
                DetectRunApp();
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

        private static void DetectRunApp(bool isConfig=false)
        {
            var getData = new ExelCreator();

            DirTemp = $@"{RootPatch}\{ConfigurationManager.AppSettings["dirTemp"]}\";
            DirEng = $@"{RootPatch}\{ConfigurationManager.AppSettings["dirEng"]}\";
            DirResRep = $@"{RootPatch}\{ConfigurationManager.AppSettings["dirResRep"]}\";
            DirCompil = $@"{RootPatch}\{ConfigurationManager.AppSettings["dirCompil"]}\";

            var dirExist = (
                Directory.Exists(DirTemp) &&
                Directory.Exists(DirEng) &&
                Directory.Exists(DirResRep) &&
                Directory.Exists(DirCompil)
                );

            if(!dirExist) { RunApp(true); return;}
            
            var temp = getData.ListTemplate(DirTemp);
            var pack = getData.ListPackage(DirEng);

            if (temp == null || pack == null)
            {
                RunApp(true);
            }
            else
            {
                if(!isConfig) UpdateSetting("rootPachExel", RootPatch);
                
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
