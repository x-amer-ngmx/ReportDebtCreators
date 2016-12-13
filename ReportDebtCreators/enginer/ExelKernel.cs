using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Authentication.ExtendedProtection;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ReportDebtCreators.model;
using Application = Microsoft.Office.Interop.Excel.Application;
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
        private Worksheet GetSheets(string shName = null, string newName = null)
        {
            /*
            var x_wBoock = _wBoock;
            var sh = (Worksheet) x_wBoock.Sheets[_wBoock.Sheets.Count];
            var rng = (Range) sh.Range["A8:AG156"];
            rng.ClearContents();
            x_wBoock.SaveAs(@"D:\popup.xlsx");*/

            if (!string.IsNullOrEmpty(newName)) _wBoock.Sheets[_wBoock.Sheets.Count].Name = newName;

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
            return _exApp.Workbooks.Open(patch,CorruptLoad:XlCorruptLoad.xlExtractData);
        }


        public void EngPackFiles(List<PackageFilesModel> packages)
        {
            
            foreach (var pack in packages)
            {
                //сводный лист по всему выбираемому диапазону
                var final = (Worksheet)_wBoock.Worksheets.Add();

                final.Name =$"Data_{pack.pack.Name}";
                var ro = 1;
                
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
                        var res =  _exApp.mRange(wS, addres);
                        ro += res.Rows.Count;
                        res.Copy(finalR);

                    }
                    catch(Exception ex)
                    {
                        var mss = ex.Message;
                        MessageBox.Show(mss);
                    }

                    wB.Close(false, Missing.Value, Missing.Value);

                    Marshal.ReleaseComObject(wS);
                    Marshal.ReleaseComObject(wB);
                }
                
            }
            //_exApp.Visible = true;

        }

        public void CreateReport(List<PackageFilesModel> packages)
        {
            //Получение из Шаблона списка Филиалов
            //Создание книг по филиалам
            //Создание листов и вставка отсортированный по филлиалу из шаблона данных
            //Сохранение книги

            GetSheets();
            //var tocopy = _wSheets;


            var one_out = false;
            foreach (var pkg in packages)
            {

                if (!one_out)
                {
                    one_out = true;
                    _wSheets.Name = pkg.pack.Name;
                }
                else
                {
                    _wSheets.Copy(After: _wSheets);
                    //_wBoock.Sheets[$"{tocopy.Name}(2)"].Name = "x";
                    GetSheets(newName: pkg.pack.Name);
                    
                }
                
                var shTmp = (Worksheet) _wBoock.Sheets.Item[$"Data_{pkg.pack.Name}"];

                var tcr = shTmp.UsedRange.Rows.Count;


                var cr = _wSheets.UsedRange.Rows.Count;

                for (var r = 7; r <= cr; r++)
                {
                    //получаем параметры поиска, по 2ум столбцам...
                    var param1 = _wSheets.Cells[r, Program.cellRange[0]].Value?.ToString();
                    var param2 = _wSheets.Cells[r, Program.cellRange[1]].Value?.ToString();

                    if ((param1 == null || param2 == null) ||
                        (string.IsNullOrEmpty(param1) || string.IsNullOrEmpty(param2))) continue;

                    if (param1.Equals("840795"))
                    {
                        var x = "stoped";
                        var mx = x;
                    }

                    Range getRow = null;

                    for (var i = 1; i < tcr; i++)
                    {
                        var tp1 = shTmp.Cells[i, 1].Value?.ToString();
                        var tp2 = shTmp.Cells[i, 2].Value?.ToString();

                        if ((tp1 != null && tp2 != null) && (param1.Equals(tp1) && param2.Equals(tp2)))
                        {
                            getRow = shTmp.UsedRange.Rows[i];
                            break;
                        }
                    }


                    //fill template

                    if (getRow == null) continue;

                    var cid = Program.cellRange.Length;

                    for (var c = 0; c < cid; c++)
                    {
                        var ci = Program.cellRange[c];
                        if (ci == Program.cellRange[0] && ci == Program.cellRange[1]) continue;
                        var v1 = getRow.Cells[1, c + 1].Value;
                        _wSheets.Cells[r, ci].Value = v1;

                    }
                }
                _exApp.DisplayAlerts = false;
                shTmp.Delete();
                _exApp.DisplayAlerts = true;

            }
            _exApp.Visible = true;

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
                SortTemplateTableInfo(p);
            }
            
        }

        private void SortTemplateTableInfo(string param)
        {
            var newWBoock = _exApp.Workbooks.Add(1);

            //var newWsheets = (Worksheet)newWBoock.ActiveSheet;


            _wSheets.Copy(newWBoock.Worksheets[1]);

            var ix = newWBoock.Worksheets.Count;
            var sh = (Worksheet) newWBoock.Worksheets[ix];

            sh.Delete();

            var newWsh = (Worksheet)newWBoock.ActiveSheet;
            
            var valid = newWsh.Range["W:W"];
            valid.Validation.Delete();

            var uses = newWsh.UsedRange;
            var uss = uses.Range[$"A7:A{uses.Rows.Count}"];

            uss.AutoFilter(23, $"<>{param}");
            uss.Delete(XlDeleteShiftDirection.xlShiftUp);

            if (newWsh.AutoFilter != null) newWsh.AutoFilterMode = false;

            
            newWsh.Protect(Password: Program.Pws,
                    DrawingObjects: _wSheets.ProtectDrawingObjects,
                    Contents:_wSheets.ProtectContents,
                    Scenarios: _wSheets.ProtectScenarios,
                    UserInterfaceOnly:_wSheets.ProtectionMode,
                    AllowFormattingCells:_wSheets.Protection.AllowFormattingCells,
                    AllowFormattingColumns: _wSheets.Protection.AllowFormattingColumns,
                    AllowFormattingRows: _wSheets.Protection.AllowFormattingRows,
                    AllowInsertingColumns: _wSheets.Protection.AllowInsertingColumns,
                    AllowInsertingRows: _wSheets.Protection.AllowInsertingRows,
                    AllowInsertingHyperlinks: _wSheets.Protection.AllowInsertingHyperlinks,
                    AllowDeletingColumns: _wSheets.Protection.AllowDeletingColumns,
                    AllowDeletingRows: _wSheets.Protection.AllowDeletingRows,
                    AllowSorting: _wSheets.Protection.AllowSorting,
                    AllowFiltering: _wSheets.Protection.AllowFiltering,
                    AllowUsingPivotTables: _wSheets.Protection.AllowUsingPivotTables);

            var f_name =
                $"{_fillPacDir}\\Перечень проблемных потребителей на {DateTime.Now:dd.MM.yyyy} {param}.xlsx";
            
            //newWBoock.Protect(Program.Pws);

            newWBoock.SaveAs(f_name, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

            newWBoock.Close(true, Missing.Value,Missing.Value);

            //Marshal.ReleaseComObject(newWsh);
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

        public static void RemoveFirstRow(this Worksheet workSheet, int x=1,int n =1)
        {
            var range = workSheet.Range[$"A{x}", $"A{n}"];
            var row = range.EntireRow;
            row.Delete(XlDirection.xlDown);
        }
        public static void RangeRemoveFirstRow(this Range rng, int n = 1)
        {
            var range = rng.Range["A1", "A" + n];
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
