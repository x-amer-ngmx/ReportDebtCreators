﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
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
                var listB = Kernel.GetListBrange("sys");

                Kernel.CreateFilseFromFill(listB);

            }
            catch (Exception ex)
            {
                var mssg = ex.Message;
                MessageBox.Show(mssg);
                //throw;
            }

            Kernel.Quit();

            Kernel = null;
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
            Kernel = new ExelKernel();

            try
            {
                // Проверка пакета файлов на совместимость шаблону

                Kernel.OpenFile(_tmpl);
                var listB = Kernel.GetListBrange("sys");
                var res = packList.EntityPackadgeFileName(listB, _form);

                if (res != null)
                {
                    Kernel.EngPackFiles(res);
                    var lst = res.OrderBy(i => i.pack.DateIndex).ToList();
                    Kernel.CreateReport(lst);
                    
                }


                //Механизм формирования отчёта для администратора
                // Определение последней книги. 
                // Определение последнего листа в книги
                // Формирование логики создания новой книги на основании последней существующей и конечного периода
                // 
            }
            catch (Exception ex)
            {
                throw;
            }

            //Kernel.Quit();

            //Kernel = null;
        }

        private void ReportForBusines(List<PackageFilesModel> packList)
        {
            Kernel = new ExelKernel();

            try
            {
                Kernel.OpenFile(_tmpl);
                var listB = Kernel.GetListBrange("sys");

                //Механизм идентификации целостности наименований файлов пакета данных.
                var res = packList.EntityPackadgeFileName(listB,_form);

                if (res != null)
                {
                    Kernel.EngPackFiles(res);
                    Kernel.CreateReport(res);
                }


                //Выделить содержимое шаблона и удалить
                // переместить подвал к заголовку, оставить 1у строку
                // И заполнить содержимое шаблона из пакета
                // сохранить как книгу и не сохранять изменений в шаблоне





                //Механизм обработки файлов пакета.

                var mmv = res;
                //Создание отчёта на основании последнего пакета данных(относительно текущей даты)
            }
            catch (Exception ex)
            {

                var ms = ex.Message;
                MessageBox.Show(ms);
                //throw;
            }


            Kernel.Quit();

            Kernel = null;
        }




    }
}
