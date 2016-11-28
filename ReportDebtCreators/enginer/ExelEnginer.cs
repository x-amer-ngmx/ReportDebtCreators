using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using ReportDebtCreators.model;
using Execl = Microsoft.Office.Interop.Excel;

namespace ReportDebtCreators.enginer
{


    /// <summary>
    /// Класс для работы с входными, выходными параметрами
    /// чистыми данными, объектами данных, реализует
    /// логику обработки данных.
    /// </summary>
    public class ExelEnginer
    {

        private readonly string _tmpl;
        private ExelKernel Kernel;
        private readonly MainCreatorsForm _form;
        public ExelEnginer()
        {
            
        }

        public ExelEnginer(string template, MainCreatorsForm form)
        {
            _tmpl = template;
            _form = form;
        }

        //основные методы которые стоит вынести за пределы исполняемого класса

        // Открытие файла,
        // Получение листов
        // Закрытие файла
        // Создание файла
        // Сохранение



        /// <summary>
        /// Открываем шаблон, сортируем данные по филлиалам, создаём файлы для заполнения филиалами
        /// </summary>
        public void CreatePackFile(string pack)
        {
            Kernel = new ExelKernel();
            try
            {
                Kernel.OpenFile(_tmpl);
                var names = Kernel.GetSheetsName();

                // var branch = Kernel.GetListBrange(names[1]);

                //var reee = branch;

                //Kernel.Quit(fname: $@"{ConfigurationManager.AppSettings["rootPachExel"]}\xui.xlsx");
                Kernel.Quit();
            }
            catch (Exception ex)
            {
                var mssg = ex.Message;
                Kernel.Quit();
                //throw;
            }

            Kernel = null;
        }


        /// <summary>
        /// Открываем шаблон, получаем список филиалов, сравниваем с колличеством файлов в пакете(спец. именованных)
        /// сравниваем структуру каждого файла, и при наличии ошибки останавливаем процесс
        /// спрашиваем пользователя окрыть ли все повреждённые файлы или просто остановить процесс.
        /// если всё впорядке то формируем отчёт.
        /// </summary>
        public void CreateReport(List<PackageFilesModel> packList, bool isRoot=false)
        {
            if (isRoot)
            {
                ReportRoot(packList);
                return;
            }

            ReportForBusines(packList);
        }


        /// <summary>
        /// Содаём отчёт для администратора системы, по выбранному пакету или диапазону пакетов
        /// где книга отчёта ведётся на протяженни 30-31 го дня с даты создания первого листа но не позднее 1го числа
        /// следующего месяца.
        /// </summary>
        private void ReportRoot(List<PackageFilesModel> packList)
        {
            Kernel = new ExelKernel();

            try
            {
                // Проверка пакета файлов на совместимость шаблону

                Kernel.OpenFile(_tmpl);
                var listB = Kernel.GetListBrange("");

                var enti = EntityPackadeFile(packList, listB);

                if(!enti) return;



                //Механизм формирования отчёта для администратора
                // Определение последней книги. 
                // Определение последнего листа в книги
                // Формирование логики создания новой книги на основании последней существующей и конечного периода
                // 
            }
            catch (Exception)
            {
                throw;
            }

            Kernel = null;
        }

        private void ReportForBusines(List<PackageFilesModel> packList)
        {
            Kernel = new ExelKernel();

            try
            {
                Kernel.OpenFile(_tmpl);
                var listB = Kernel.GetListBrange("");

                var enti = EntityPackadeFile(packList, listB);

                if (!enti) return;


                //Создание отчёта на основании последнего пакета данных(относительно текущей даты)
            }
            catch (Exception)
            {
                throw;
            }

            Kernel = null;
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
        public List<PackageFilesModel> GetEngineFList(IList<StructExelModel> packRange)
        {
            if (packRange == null || !packRange.Any()) return null;

            var result = new List<PackageFilesModel>();
            result.AddRange(
                from pack in packRange
                let file = new DirectoryInfo(pack.AbsolutPatch).GetFilesByExtensions(".xlsx", ".xls")
                select new PackageFilesModel
                {
                    pack = pack,
                    BrangeFiles = (from t in file select new StructExelModel { Name = t.Name.Split('.')[0], AbsolutPatch = t.FullName }).ToList()
                });

            return result;
        }


        private bool EntityPackadeFile(List<PackageFilesModel> packList, List<string> listB)
        {
            foreach (var reswin in
            from pack in packList
            let cbn = pack.BrangeFiles.Count
            let mxx = pack.BrangeFiles.Sum(brn => (
            from br in listB
            let reg = new Regex($".*{br}.*", RegexOptions.IgnoreCase)
            select reg.IsMatch(brn.Name) ? 1 : 0).Sum())
            where mxx < cbn || mxx > cbn
            select MessageBox.Show($"{(cbn - mxx)}", "", MessageBoxButtons.YesNo, MessageBoxIcon.Information))
            {
                if (reswin == DialogResult.Yes)
                {
                    var m = MainCreatorsForm.ActiveForm;
                    
                    //Вывод предупреждения о том что пакет был не полон или
                    // некоторые файлы были именованны не корректно.
                    // 

                _form.SetInfoLable("Всё плохо но мы продолжаем!");
                }
                else
                {
                    _form.SetInfoLable("Всё плохо и мы останавливаемся!");
                    //Выход из метода и вывод предупреждения.

                }
            }
            return false;
        }


    }
}
