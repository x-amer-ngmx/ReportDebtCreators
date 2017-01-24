using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using ReportDebtCreators.model;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ReportDebtCreators.enginer
{
    /// <summary>
    /// Класс для работы с логикой экселя
    /// </summary>
    class ExelKernel
    {
        private readonly Application _exApp;
        private Workbook _wBoock;
        private Worksheet _wSheets;

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
        /// <param name="newName"></param>
        /// <returns></returns>
        private Worksheet GetSheets(string shName = null, string newName = null)
        {
            if (!string.IsNullOrEmpty(newName)) _wBoock.Sheets[_wBoock.Sheets.Count].Name = newName;

            //получаем последний лист из книги (это очень важно)
            _wSheets = string.IsNullOrEmpty(shName) ? (Worksheet)_wBoock.Sheets[_wBoock.Sheets.Count] : (Worksheet) _wBoock.Sheets.Item[shName];
            _wSheets.Unprotect(Program.Pws); //снимаем защиту.
            return _wSheets;
        }

        public void EngPackFiles(List<PackageFilesModel> packages)
        {
            if (File.Exists(Program.TempJsonDB)) File.Delete(Program.TempJsonDB);

            var cl = GetColName(Program.cellRange[Program.cellRange.Count() - 1]);

            var itemsPack = new Dictionary<string, object>();
            foreach (var pack in packages)
            {

                var listItemsRow= new Dictionary<string, object>();
                foreach (var pm in pack.BrangeFiles)
                {
                    using (var conn = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{pm.AbsolutPatch}';Extended Properties=\"Excel 12.0 Xml;HDR=No;IMEX=0;MAXSCANROWS=6;READONLY=FALSE\";"))
                    {
                        conn.Open();
                        var sh = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        if(sh==null) continue;
                        var cmd = conn.CreateCommand();
                        var str = $"SELECT * FROM [{sh.Rows[0]["TABLE_NAME"].ToString().Replace("'",string.Empty)}A6:{cl}]";
                        cmd.CommandText = str;
                        
                        using (var rdr = cmd.ExecuteReader())
                        {

                            var query =
                                (from DbDataRecord row in rdr select row);

                            var litem = (from dbRw in query
                                where
                                    !string.IsNullOrEmpty(dbRw[Program.cellRange[0]].ToString()) &&
                                    !string.IsNullOrEmpty(dbRw[Program.cellRange[1]].ToString())
                                select Program.cellRange.ToDictionary(i => $"F{i}", i => dbRw[i-1])).ToList();

                            
                            listItemsRow.Add(pm.Name,litem);

                        }
                    }
                }
                itemsPack.Add(pack.pack.Name, listItemsRow);
            }

            var json = JsonConvert.SerializeObject(itemsPack);
            File.WriteAllText(Program.TempJsonDB, json);

        }

        public void CreateReport(List<PackageFilesModel> packages)
        {
            GetSheets();

            var oneOut = false;

            var openFile = File.ReadAllText(Program.TempJsonDB);
            var data = JsonConvert.DeserializeObject<Dictionary<string,object>>(openFile);

            foreach (var pack in packages)
            {

                if (!oneOut)
                {
                    oneOut = true;
                    _wSheets.Name = pack.pack.Name;
                }
                else
                {
                    _wSheets.Copy(After: _wSheets);
                    GetSheets(newName: pack.pack.Name);

                }

                var cr = _wSheets.UsedRange.Rows.Count;

                var getPak =
                    (from x in data
                        where x.Key.Equals(pack.pack.Name)
                        select JsonConvert.DeserializeObject<Dictionary<string, object>>(x.Value.ToString())).Single();

                var row = pack.BrangeFiles.Select(brn => (from gg in getPak
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

                _wSheets.Range["D4"].Select();
            }

            var fileX = $"{Program.DirResRep}Отчёт на {DateTime.Now:dd.MM.yyyy hh_mm_ss}.xlsx";
            _wBoock.SaveAs(fileX, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);
        }


        private void AddThesData(List<List<Dictionary<string, object>>> row, int r, string param1, string param2)
        {
            foreach (var rowList in row)
            {
                foreach (var xRow in rowList)
                {
                    var p1 = (xRow.Values.ElementAt(0) ?? "").ToString();
                    var p2 = (xRow.Values.ElementAt(1) ?? "").ToString();

                    if (!p1.Equals(param1) || !p2.Equals(param2)) continue;

                    foreach (var clr in Program.cellRange.Skip(2))
                    {
                        var frm = _wSheets.Cells[r, clr];
                        var formul = ((Range)frm).Formula.ToString();

                        var reg = new Regex("^=.*", RegexOptions.IgnoreCase);

                        if (reg.IsMatch(formul))
                        {
                            continue;
                        }
                        var v1 = xRow[$"F{clr}"];
                        frm.Value = v1;
                    }
                }
            }
        }


        public void CreateFilseFromFill(IEnumerable<string> param)
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

            var fName =
                $"{_fillPacDir}\\Перечень проблемных потребителей на {DateTime.Now:dd.MM.yyyy} {param}.xlsx";
            
            //newWBoock.Protect(Program.Pws);
            //newWsh.SaveAs(f_name, Type.Missing, Program.Pws, Type.Missing, false,false);
            newWBoock.SaveAs(fName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);
            
            newWBoock.Close(true, Missing.Value,Missing.Value);

            Marshal.ReleaseComObject(newWsh);
            Marshal.ReleaseComObject(newWBoock);
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
        public static int ColumnNameToNumber(this string colName)
        {
            var result = 0;

            // Process each letter.
            foreach (var t in colName)
            {
                result *= 26;
                var letter = t;

                // See if it's out of bounds.
                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                // Add in the value of this letter.
                result += letter - 'A' + 1;
            }
            return result;
        }
    }
}
