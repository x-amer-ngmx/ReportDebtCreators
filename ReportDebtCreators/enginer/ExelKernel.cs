using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using ReportDebtCreators.model;
using DataTable = System.Data.DataTable;

namespace ReportDebtCreators.enginer
{
    /// <summary>
    /// Класс для работы с логикой экселя
    /// </summary>
    class ExelKernel
    {
        private Application _exApp;
        private Workbook _wBoock;
        private Worksheet _wSheets;
        private Sheets _sheets;

        private string _fillPacDir;

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
           _wBoock = _exApp.Workbooks.Open(patch);
        }


        public List<string> GetSheetsName()
        {
            return (from Worksheet sh in _wBoock.Sheets select sh.Name).ToList();
        }

        public List<string> GetListBrange(string name)
        {
            var sh = GetSheets(name);
            var ur = sh.UsedRange;
            var arr = (Range)ur.Rows[$"2:{ur.Rows.Count}"];
            var result = (from Range x in arr where x.Value != null select (string)x.Value).ToList();
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
            if(_wSheets != null) Marshal.ReleaseComObject(_wSheets);
            if(_sheets != null) Marshal.ReleaseComObject(_sheets);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        #region Привытные методы

        /// <summary>
        /// Метод возвращает последний лист или лист по имени
        /// </summary>
        /// <param name="shName">если наименование листа не заданно то метод вернёт последний</param>
        /// <returns></returns>
        private Worksheet GetSheets(string shName = null)
        {
            /*
            var x_wBoock = _wBoock;
            var sh = (Worksheet) x_wBoock.Sheets[_wBoock.Sheets.Count];
            var rng = (Range) sh.Range["A8:AG156"];
            rng.ClearContents();
            x_wBoock.SaveAs(@"D:\popup.xlsx");*/

            //получаем последний лист из книги (это очень важно)
            _wSheets = string.IsNullOrEmpty(shName) ? (Worksheet)_wBoock.Sheets[_wBoock.Sheets.Count] : (Worksheet) _wBoock.Sheets.Item[shName];
            //_wSheets.Protect(Program.Pws);
            _wSheets.Unprotect(Program.Pws); //снимаем защиту.
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

        private _Workbook OpenPackFile(string patch)
        {
            return _exApp.Workbooks.Open(patch);
        }

        public void EngPackFiles(List<PackageFilesModel> packages)
        {
            //сводный лист по всему выбираемому диапазону
            var final = (Worksheet)_wBoock.Worksheets.Add();
            var ro = 1;
            
            foreach (var pack in packages)
            {
                foreach (var pm in pack.BrangeFiles)
                {
                    var finalR = final.Range[$"A{ro}"];

                    var wB = OpenPackFile(pm.AbsolutPatch);
                    var wS = (_Worksheet)wB.ActiveSheet;

                    //выделяем диапазон необходимых данных
                    //стольбец лицевых и заполняемых значений 
                    // (далее по лицевым будем заполнять шаблон)
                    //далее по логике построения отчёта
                    //вставляем в шаблон
                    //Сохраняем отчёт

                    //Исключить шапку и подвал документа
                    var usR = wS.UsedRange;
                    var rc = usR.Rows.Count;

                    var addres = GetAddrRange(6, rc, Program.cellRange);

                    
                    try
                    {

                        //создаём временный лист
                        var x = (Worksheet)_wBoock.Worksheets.Add();
                        var r = x.Range["A1"];
                        //получаем данные из вычисляемых столбцов
                        var res = _exApp.mRange(wS, addres);
                        //помещаем в него результаты по всем вычесляемым столбцам
                        res.Copy(r);
                        
                        //удаляем заголовки
                        x.RemoveFirstRow();

                        //определяем область копирования
                        var m = x.UsedRange;

                        //расчитываем следующую позицию для вставки блока данных
                        ro += m.Rows.Count;

                        //вставляем данные в накопительный лист
                        m.Copy(finalR);

                        //удаляем промежуточный временный лист
                        _exApp.DisplayAlerts = false;
                        x.Delete();
                        _exApp.DisplayAlerts = true;
                    }
                    catch(Exception ex)
                    {
                        var mss = ex.Message;
                    }

                    //var rWS = wS.Range["E6:E126,G:G,H:H,I:I,J:J,N:N,V:V,X:X"];

                    wB.Close(false, Missing.Value, Missing.Value);

                    Marshal.ReleaseComObject(wS);
                    Marshal.ReleaseComObject(wB);
                }
                
            }
/*
            var x = (_Worksheet)_wBoock.Worksheets.Add();
            
            foreach (var re in res)
            {
               var r = x.Range[$"A{(re.Rows.Count+1)}"];
                re.Copy(r);
            }*/
            


 _exApp.Visible = true;
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

        public void CreateFilseFromFill(List<string> param)
        {
            _fillPacDir = $"{Program.DirCompil}Для филиалов за {DateTime.Now:dd.MM.yyyy}";
            if (!Directory.Exists(_fillPacDir)) Directory.CreateDirectory(_fillPacDir);

            GetSheets();
            foreach (var p in param)
            {
                SortTemplateTableInfo(_wSheets, p);
            }
            
        }

        private void SortTemplateTableInfo(Worksheet ws, string param)
        {
            var rangeList = (Range)ws.UsedRange;

            
            rangeList.AutoFilter(23, param);

            var filtered = rangeList.SpecialCells(XlCellType.xlCellTypeVisible);

            var newWBoock = (_Workbook)_exApp.Workbooks.Add();
            var newWsheets = (_Worksheet)newWBoock.ActiveSheet;

            var nran = newWsheets.Range["A1"];

            filtered.Copy(nran);


            var f_name =
                $"{_fillPacDir}\\Перечень проблемных потребителей на {DateTime.Now:dd.MM.yyyy} {param}.xlsx";

            newWBoock.SaveAs(f_name, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

            newWBoock.Close(true, Missing.Value,Missing.Value);

            Marshal.ReleaseComObject(newWsheets);
            Marshal.ReleaseComObject(newWBoock);
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

        private string GetAddrRange(int rowBegin, int rowEnd, int[] cellIndexRange)
        {
            var address = string.Empty;
            foreach (var i in cellIndexRange)
            {
                var cN = GetColName(i);
                address += $"{cN}{rowBegin}:{cN}{rowEnd}{(i!= cellIndexRange[cellIndexRange.Count() - 1] ? ",":string.Empty)}";

            }
            // A1:A99,D1:D99 ...

            return address;
        }

        private string GetColName(int index)
        {
            var range = "";
            index -= 1;
            if (index < 1) return range;
            for (var i = 1; index + i > 0; i = 0)
            {
                range = $"{((char)(65 + index % 26))}{range}";
                index /= 26;
            }
            if (range.Length > 1) range = $"{((char)((int)range[0] - 1))}{range.Substring(1)}";
            return range;
        }

        #endregion

    }


    static class Custommer
    {
        public static Range mRange(this Application app, _Worksheet ws, string addr)
        {
            var address = addr.Split(',');
            return address.Select(s => ws.Range[s]).Aggregate<Range, Range>(null, app.Unionx);
        }

        public static Range Unionx(this Application app, Range r0, Range r1)
        {
            if (r0 == null && r1 == null)
                return null;
            if (r0 == null)
                return r1;
            if (r1 == null)
                return r0;
            return app.Union(r0, r1);
        }

        public static void RemoveFirstRow(this Worksheet workSheet, int n =1)
        {
            var range = workSheet.Range["A1", "A" + n];
            var row = range.EntireRow;
            row.Delete(XlDirection.xlUp);
        }

        public static DataTable GetDataTable(this Range rn)
        {
            var tbl = new DataTable ();

            var cc = rn.Columns.Count;
            var rc = rn.Rows.Count;

            for (var i = 0; i < cc; i++)
            {
                 var x = (object)rn.Cells[1, i + 1].Value2;
                tbl.Columns.Add(x.ToString(),typeof(string));
            }

            for (var i = 0; i < cc; i++)
            {
                
                for (var j = 0; j < rc; j++)
                {
                   

                }
                
            }

            return tbl;
        }
    }
}
