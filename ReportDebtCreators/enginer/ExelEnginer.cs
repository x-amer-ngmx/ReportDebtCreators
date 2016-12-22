using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ReportDebtCreators.model;

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
        private ExelKernel _kernel;
        private readonly MainCreatorsForm _form;

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
        public void CreatePackFile()
        {
            _kernel = new ExelKernel();
            try
            {
                _kernel.OpenFile(_tmpl);
                var listB = _kernel.GetListBrange("sys");
                _kernel.CreateFilseFromFill(listB);
            }
            catch (Exception ex)
            {
                var mssg = ex.Message;
                MessageBox.Show(mssg);
            }

            _kernel.Quit();
            _kernel = null;
        }


        /// <summary>
        /// Открываем шаблон, получаем список филиалов, сравниваем с колличеством файлов в пакете(спец. именованных)
        /// сравниваем структуру каждого файла, и при наличии ошибки останавливаем процесс
        /// спрашиваем пользователя окрыть ли все повреждённые файлы или просто остановить процесс.
        /// если всё впорядке то формируем отчёт.
        /// </summary>
        public void CreateReport(List<PackageFilesModel> packList, bool isAdmin=false)
        {
            if (isAdmin)
            {
                ReportAdmin(packList);
                return;
            }

            ReportForBusines(packList);
        }


        /// <summary>
        /// Содаём отчёт для администратора системы, по выбранному пакету или диапазону пакетов
        /// где книга отчёта ведётся на протяженни 30-31 го дня с даты создания первого листа но не позднее 1го числа
        /// следующего месяца.
        /// </summary>
        private void ReportAdmin(List<PackageFilesModel> packList)
        {
            _kernel = new ExelKernel();

            try
            {
                // Проверка пакета файлов на совместимость шаблону

                _kernel.OpenFile(_tmpl);
                var listB = _kernel.GetListBrange("sys");
                var res = packList.EntityPackadgeFileName(listB, _form);

                if (res != null)
                {
                    _kernel.EngPackFiles(res);
                    var lst = res.OrderBy(i => i.pack.DateIndex).ToList();
                    _kernel.CreateReport(lst);
                }


                //Механизм формирования отчёта для администратора
                // Определение последней книги. 
                // Определение последнего листа в книги
                // Формирование логики создания новой книги на основании последней существующей и конечного периода
                // 
            }
            catch (Exception ex)
            {
                var ms = ex.Message;
                MessageBox.Show(ms);
            }

            _kernel.Quit();
            _kernel = null;
        }

        private void ReportForBusines(List<PackageFilesModel> packList)
        {
            _kernel = new ExelKernel();

            try
            {
                _kernel.OpenFile(_tmpl);
                var listB = _kernel.GetListBrange("sys");

                //Механизм идентификации целостности наименований файлов пакета данных.
                var res = packList.EntityPackadgeFileName(listB,_form);

                if (res != null)
                {
                    _kernel.EngPackFiles(res);
                    _kernel.CreateReport(res);
                }


                //Выделить содержимое шаблона и удалить
                // переместить подвал к заголовку, оставить 1у строку
                // И заполнить содержимое шаблона из пакета
                // сохранить как книгу и не сохранять изменений в шаблоне
                //Механизм обработки файлов пакета.
                //Создание отчёта на основании последнего пакета данных(относительно текущей даты)
            }
            catch (Exception ex)
            {
                var ms = ex.Message;
                MessageBox.Show(ms);
            }

            _kernel.Quit();
            _kernel = null;
        }
    }
}
