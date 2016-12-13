using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ReportDebtCreators.model;

namespace ReportDebtCreators.enginer
{
    public static class Global
    {
        public static IEnumerable<FileInfo> GetFilesByExtensions(this DirectoryInfo dir, params string[] extensions)
        {
            if (extensions == null)
                throw new ArgumentNullException("extensions");
            IEnumerable<FileInfo> files = dir.EnumerateFiles();
            return files.Where(f => extensions.Contains(f.Extension));
        }

        public static IList<StructExelModel> GetRangePack(this IList<StructExelModel> fromPack,DateTime? f, DateTime? t)
        {
            return (from p in fromPack
             orderby p.DateIndex descending
             where (p.DateIndex >= f && p.DateIndex <= t)
             select p).ToList();
        }

        public static IList<StructExelModel> GetPack(this IList<StructExelModel> fromPack, DateTime? pName)
        {
            return (from p in fromPack
                    orderby p.DateIndex descending
                    where (p.DateIndex == pName)
                    select p).ToList();
        }

        /// <summary>
        /// Получаем список файлов в пакете или диапазоне пакетов
        /// для формирования отчётов
        /// в формате :
        /// /пакет
        /// /--файл1
        /// /--файл2
        /// /-- ...
        /// /--файлN
        /// 
        /// полученные данные будут в последствии обработанны
        /// на определение полноты пакета.
        /// </summary>
        /// <returns></returns>
        public static List<PackageFilesModel> GetEngineFList(this IList<StructExelModel> packRange)
        {
            if (packRange == null || !packRange.Any()) return null;

            var result = new List<PackageFilesModel>();
            result.AddRange(
                from pack in packRange
                let file = new DirectoryInfo(pack.AbsolutPatch).GetFilesByExtensions(".xlsx", ".xls")
                select new PackageFilesModel
                {
                    pack = pack,
                    BrangeFiles = (from t in file select new StructExelModel { Name = t.Name.Substring(0, t.Name.LastIndexOf('.')), AbsolutPatch = t.FullName }).ToList()
                });

            return result;
        }

        public static List<PackageFilesModel> EntityPackadgeFileName(this List<PackageFilesModel> packList, List<string> listB, MainCreatorsForm form)
        {
            var cB = listB.Count;
            var bout = false;
            List<StructExelModel> rnfr = new List<StructExelModel>();
            List<PackageFilesModel> rrs = null;

           foreach (var pk in packList)
            {

                var cpkb = 0;
                foreach (var pkb in pk.BrangeFiles)
                {
                    var id =
                        listB.Select(brx => new Regex($@"^([^\$\~]).*{brx}.*", RegexOptions.IgnoreCase))
                            .Count(reg => reg.IsMatch(pkb.Name));
                    cpkb += id;

                    if (id == 0)
                    {
                        rnfr.Add(pkb);
                    }
                }

                if (cpkb < cB || cpkb > cB)
                {
                    var win = MessageBox.Show($"{(cB - cpkb)}", "", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                    if (win == DialogResult.Yes)
                    {
                        form.SetInfoLable("Всё плохо но мы продолжаем!");
                    }
                    else
                    {
                        form.SetInfoLable("Всё плохо и мы останавливаемся!");
                        break;
                    }
                }

                if (rnfr.Any())
                {

                    foreach (var xx in rnfr)
                    {
                        pk.BrangeFiles.Remove(xx);
                    }
                }

                if (pk.BrangeFiles.Any())
                {
                    rrs = rrs ?? new List<PackageFilesModel>();
                    rrs.Add(pk);
                }
                
                //^
            }

            return rrs;
            /*
            foreach (var reswin in
            from pack in packList
                let cbn = listB.Count
                let mxx = pack.BrangeFiles.Sum(brn => (
                    from br in listB
                        let reg = new Regex($".*{br}.*", RegexOptions.IgnoreCase)
                        select reg.IsMatch(brn.Name) ? 1 : 0).Sum())
                where mxx < cbn || mxx > cbn
            select MessageBox.Show($"{(cbn - mxx)}", "", MessageBoxButtons.YesNo, MessageBoxIcon.Information))
            {
                if (reswin == DialogResult.Yes)
                {
                    //Вывод предупреждения о том что пакет был не полон или
                    // некоторые файлы были именованны не корректно.
                    // 
                    result = true;
                    _form.SetInfoLable("Всё плохо но мы продолжаем!");
                }
                else
                {
                    _form.SetInfoLable("Всё плохо и мы останавливаемся!");
                    return result;
                    //Выход из метода и вывод предупреждения.

                }
            }
            */
        }

        //Анализатор имён директорий, 
        public static DateTime? PacNameConvert(this string packName)
        {
            DateTime? result = null;
            try
            {
                result = DateTime.Parse(packName);
            }
            catch (Exception)
            {

                //throw;
            }


            return result;
        }

        //Переводим Range в DataTable, для последующей обработки данных
  }
}
