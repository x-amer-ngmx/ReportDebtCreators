using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ReportDebtCreators.enginer
{
    class ExelKernel
    {
        private Application _exApp;
        private Workbook _wBoock;
        private Workbooks _wBoocks;
        private Worksheet _wSheets;
        private Sheets _sheets;

        public ExelKernel()
        {
            _exApp = new Application();
        }


        //Глобальные методы

        //Открыть файл
        //Получить содержимое
        //Методы обработки содержимого
        //Закрыть файл(сохранить/не сохранять)


        #region Публичные методы
        public void OpenFile(string patch)
        {
           _wBoocks = _exApp.Workbooks;
           _wBoock = _wBoocks.Open(patch);
        }

        public List<string> GetSheetsName()
        {
            return (from Worksheet sh in _wBoock.Sheets select sh.Name).ToList();
        }

        public void Quit(bool save=false, string fname=null)
        {
            if(save)_wBoock.Save();
            if(fname!=null) _wBoock.SaveAs(fname);
            _wBoock.Close(0);
            _exApp.Quit();

            Marshal.ReleaseComObject(_exApp);
            Marshal.ReleaseComObject(_wBoock);
            Marshal.ReleaseComObject(_wBoocks);
           /* Marshal.ReleaseComObject(WSheets);
            Marshal.ReleaseComObject(TSheets);*/

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        #region Привытные методы

        private Worksheet GetSheets(string shName)
        {
            _wSheets = (Worksheet)_wBoock.Sheets.Item[shName];
            _wSheets.Protect(Program.Pws);
            return _wSheets;
        }


        private void CreatePackage()
        {
            //Получение из Шаблона списка Филиалов
            //Создание книг по филиалам
            //Создание листов и вставка отсортированный по филлиалу из шаблона данных
            //Сохранение книги
        }

        private void CreateRootReport()
        {
            var wb = CreateWorkbook();
            //Создание новой книги и листа
            //Вставка отсортированных данных из шаблона
            //Сохранение отчёта

        }

        private void CreateAdminReport()
        {

            //Поиск последнего файла
            //Получение диапазона имён листов
            //сравнение с добавляемой информацией
            //добавление данных / либо создание новой книги

        }

        private void SortTemplateTableInfo(Worksheet ws)
        {
            //Сортировка данных из Шаблона
        }

        private void CreateWorksheet(ref Workbook wb, Worksheet ws)
        {
            wb.Worksheets.Add(ws);
        }

        private Workbook CreateWorkbook()
        {
            return _exApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            //return (Worksheet) wb.Worksheets[DateTime.Now.ToString("dd.MM.yyyy")];
        }



        #endregion

    }
}
