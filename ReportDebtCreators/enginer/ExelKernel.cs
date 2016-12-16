using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Authentication.ExtendedProtection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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
            if (File.Exists(Program.TempJsonDB)) File.Delete(Program.TempJsonDB);

            var ItemsPack = new Dictionary<string, object>();
            foreach (var pack in packages)
            {

                //сводный лист по всему выбираемому диапазону
               /* var final = (Worksheet)_wBoock.Worksheets.Add();

                final.Name =$"Data_{pack.pack.Name}";
                var ro = 1;*/

                var ListItemsRow= new Dictionary<string, object>();
                foreach (var pm in pack.BrangeFiles)
                {
                    // var finalR = final.Range[$"A{ro}"];


                    using (var conn = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{pm.AbsolutPatch}';Extended Properties=\"Excel 12.0;HDR=YES;\";"))
                    {
                        conn.Open();
                        var sh = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                        var cmd = conn.CreateCommand();
                        cmd.CommandText = $"SELECT * FROM [{sh.Rows[0]["TABLE_NAME"]}] ";

                        using (var rdr = cmd.ExecuteReader())
                        {
                            var query =
                                (from DbDataRecord row in rdr select row);



                            var Litem = (from dbRw in query
                                where
                                    !string.IsNullOrEmpty(dbRw[Program.cellRange[0]].ToString()) &&
                                    !string.IsNullOrEmpty(dbRw[Program.cellRange[1]].ToString())
                                select Program.cellRange.ToDictionary(i => $"F{i}", i => dbRw[i-1])).ToList();

                            var dic = new Dictionary<string,object>() {};
                            
                            ListItemsRow.Add(pm.Name,Litem);

                        }
                    }

                    /*
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



                    
                   /* 
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
                    }*

                    wB.Close(false, Missing.Value, Missing.Value);

                    Marshal.ReleaseComObject(wS);
                    Marshal.ReleaseComObject(wB);*/
                }

                ItemsPack.Add(pack.pack.Name, ListItemsRow);
            }

            var json = JsonConvert.SerializeObject(ItemsPack);
            //_exApp.Visible = true;

            File.WriteAllText(Program.TempJsonDB, json);

        }

        public void CreateReport(List<PackageFilesModel> packages)
        {
            GetSheets();
            //var tocopy = _wSheets;

            var one_out = false;


            var open_file = File.ReadAllText(Program.TempJsonDB);
            var data = JsonConvert.DeserializeObject<Dictionary<string,object>>(open_file);

            foreach (var pack in packages)
            {

                if (!one_out)
                {
                    one_out = true;
                    _wSheets.Name = pack.pack.Name;
                }
                else
                {
                    _wSheets.Copy(After: _wSheets);
                    GetSheets(newName: pack.pack.Name);

                }

                var cr = _wSheets.UsedRange.Rows.Count;

                var get_pak =
                    (from x in data
                        where x.Key.Equals(pack.pack.Name)
                        select JsonConvert.DeserializeObject<Dictionary<string, object>>(x.Value.ToString())).Single();

                var row = pack.BrangeFiles.Select(brn => (from gg in get_pak
                    where gg.Key.Equals(brn.Name)
                    select JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(gg.Value.ToString())).Single())
                    .ToList();


                for (var r = 7; r <= cr; r++)
                {
                    //получаем параметры поиска, по 2ум столбцам...
                    var param1 = _wSheets.Cells[r, Program.cellRange[0]].Value?.ToString();
                    var param2 = _wSheets.Cells[r, Program.cellRange[1]].Value?.ToString();

                    if ((param1 == null || param2 == null) ||
                        (string.IsNullOrEmpty(param1) || string.IsNullOrEmpty(param2))) continue;

                    AddThesData(row, r, param1, param2);

                }
            }

            _exApp.Visible = true;
        }


        private void AddThesData(List<List<Dictionary<string, object>>> row, int r, string param1, string param2)
        {
            foreach (var _rowList in row)
            {
                foreach (var _row in _rowList)
                {
                    var p1 = (_row.Values.ElementAt(0) ?? "").ToString();
                    var p2 = (_row.Values.ElementAt(1) ?? "").ToString();

                    if (p1.Equals(param1) && p2.Equals(param2))
                    {

                        foreach (var clr in Program.cellRange.Skip(2))
                        {
                            var frm = _wSheets.Cells[r, clr];
                            var formul = ((Range)frm).Formula.ToString();

                            var reg = new Regex("^=.*", RegexOptions.IgnoreCase);

                            if (reg.IsMatch(formul))
                            {
                                continue;
                            }
                            var v1 = _row[$"F{clr}"];
                            frm.Value = v1;
                        }

                        return;
                    }
                }
            }
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

            var valid = newWsh.Range[$"{Program.brnCell}:{Program.brnCell}"];
            valid.Validation.Delete();

            var uses = newWsh.UsedRange;
            var rc = uses.Rows.Count;
            var uss = uses.Range[$"A7:A{rc}"];

            uss.AutoFilter(Program.brnCell.ColumnNameToNumber(), $"<>{param}");
            uss.Delete(XlDeleteShiftDirection.xlShiftUp);

            if (newWsh.AutoFilter != null) newWsh.AutoFilterMode = false;


            var xuses = newWsh.UsedRange;
            var xrc = xuses.Rows.Count+2;

            var rng = xuses.Range[$"A{xrc}:A{xrc + 4}"];
            var rrng = rng.EntireRow;
            rrng.Locked = false;


            newWsh.Range["E7"].Select();
            newWsh.Protect(Password: Program.Pws,
                    DrawingObjects: _wSheets.ProtectDrawingObjects,
                    Contents:true,
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
                    AllowFiltering: false,
                    AllowUsingPivotTables: _wSheets.Protection.AllowUsingPivotTables);

            var f_name =
                $"{_fillPacDir}\\Перечень проблемных потребителей на {DateTime.Now:dd.MM.yyyy} {param}.xlsx";
            
            //newWBoock.Protect(Program.Pws);
            //newWsh.SaveAs(f_name, Type.Missing, Program.Pws, Type.Missing, false,false);
            newWBoock.SaveAs(f_name, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);
            
            newWBoock.Close(true, Missing.Value,Missing.Value);

            Marshal.ReleaseComObject(newWsh);
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


        public static int ColumnNameToNumber(this string col_name)
        {
            int result = 0;

            // Process each letter.
            for (int i = 0; i < col_name.Length; i++)
            {
                result *= 26;
                char letter = col_name[i];

                // See if it's out of bounds.
                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                // Add in the value of this letter.
                result += (int)letter - (int)'A' + 1;
            }
            return result;
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
