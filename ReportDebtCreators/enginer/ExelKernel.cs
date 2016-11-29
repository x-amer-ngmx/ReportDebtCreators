using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ReportDebtCreators.enginer
{
    /// <summary>
    /// Класс для работы с логикой экселя
    /// </summary>
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

        public List<string> GetListBrange(string name)
        {
            var sh = GetSheets(name);
            var res = GetMaxMinWorcRange(sh);
            var arr = sh.Range[res];

            var result = (from Range x in arr.Cells where x.Value != null select (string)x.Value).ToList();

            return result;
        }

        public void Quit(bool save=false, string fname=null)
        {
            _wSheets?.Protect(Program.Pws);
            if(save)_wBoock.Save();
            if(fname!=null) _wBoock.SaveAs(fname);
            _wBoock?.Close(0);
            _exApp.Quit();

            Marshal.ReleaseComObject(_exApp);
            if(_wBoock != null) Marshal.ReleaseComObject(_wBoock);
            if(_wBoocks != null) Marshal.ReleaseComObject(_wBoocks);
            if(_wSheets != null) Marshal.ReleaseComObject(_wSheets);
            if(_sheets != null) Marshal.ReleaseComObject(_sheets);

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
            //_wSheets.Protect(Program.Pws);
            _wSheets.Unprotect(Program.Pws);
            return _wSheets;
        }


        /*
        private List<string> GetRangeBranch(_Worksheet xSheet)
        {
            var result = new List<string>();
            try
            {
                var row = xSheet.UsedRange.Rows.Count;
                var cel = xSheet.UsedRange.Columns.Count;

                for (var i = 1; i <= row; i++)
                {
                    for (var j = 1; j <= cel; j++)
                    {
                        var vall = (Range) xSheet.Cells[i, j];

                        if (vall != null && vall.Value != null && string.IsNullOrEmpty(vall.Value.ToString()))
                        {
                            result.Add($"{vall.Value}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //throw;
            }

            return result;
        }
        */


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



        private string GetMaxMinWorcRange(_Worksheet xSheet)
        {
            var result = string.Empty;
            try
            {
                var rowMax = 0;
                var colMax = 0;

                var usedRange = xSheet.UsedRange;
                var lastCell = usedRange.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                var lastRow = lastCell.Row;
                var lastCol = lastCell.Column;
                var rowMin = lastRow + 1;
                var colMin = lastCol + 1;

                var row = xSheet.UsedRange.Rows.Count;
                var cel = xSheet.UsedRange.Columns.Count;

                for (var i = 1; i <= row; i++)
                {
                    for (var j = 1; j <= cel; j++)
                    {
                        var vall = (Range)xSheet.Cells[i, j];

                        if (vall != null && vall.Value != null && !string.IsNullOrEmpty(vall.Value.ToString()))
                        {
                            if (vall.Row > rowMax)
                                rowMax = vall.Row;

                            if (vall.Column > colMax)
                                colMax = vall.Column;

                            if (vall.Row < rowMin)
                                rowMin = vall.Row+1;

                            if (vall.Column < colMin)
                                colMin = vall.Column;
                        }
                        MRCO(vall);
                    }
                }


                if (!(rowMax == 0 || colMax == 0 || rowMin == lastRow + 1 || colMin == lastCol + 1))
                    result = Cells2Address(rowMin, colMin, rowMax, colMax);

                MRCO(lastCell);
                MRCO(usedRange);

            }
            catch (Exception ex)
            {
                //throw;
            }

            return result;
        }

        private string Cells2Address(int row1, int col1, int row2, int col2)
        {
            return $"{ColNum2Letter(col1)}{row1}:{ColNum2Letter(col2)}{row2}";
        }

        private string ColNum2Letter(int colNum)
        {
            if (colNum <= 26)
                return ((char)(colNum + 64)).ToString();

            colNum--; //decrement to put value on zero based index
            return ColNum2Letter(colNum / 26) + ColNum2Letter((colNum % 26) + 1);
        }

        private void MRCO(object obj)
        {
            if (obj == null) { return; }
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch
            {
                // ignore, cf: http://support.microsoft.com/default.aspx/kb/317109
            }
            finally
            {
                obj = null;
            }
        }

        #endregion

    }
}
