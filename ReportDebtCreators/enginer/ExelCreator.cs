using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using ReportDebtCreators.model;


namespace ReportDebtCreators.enginer
{
    class ExelCreator
    {
        //Получение путей к конечным файлам

        public List<StructExelModel> ListTemplate(string dir)
        {
            List<StructExelModel> result = null;
            try
            {
                var file = new DirectoryInfo(dir).GetFilesByExtensions(".xltx", ".xlt");
                result = (from t in file select new StructExelModel {Name = t.Name.Split('.')[0], AbsolutPatch = t.FullName}).ToList();
            }
            catch (Exception ex)
            {
                // ignored
            }


            return result;
        }

        public List<StructExelModel> ListPackage(string patch)
        {
            List<StructExelModel> result = null;

            try
            {
                var file = new DirectoryInfo(patch).GetDirectories("*");
                result = (from t in file
                          let id = t.Name.PacNameConvert()
                          where id != null
                          orderby id descending
                          select new StructExelModel { Name = t.Name, AbsolutPatch = $"{t.FullName}\\", DateIndex = id }).ToList();


            }
            catch (Exception ex)
            {
                // ignored
            }

            return result;
        }

        public static List<StructExelModel> ListPackageFiles(string patch)
        {
            List<StructExelModel> result = null;

            try
            {
                var files = new DirectoryInfo(patch).GetFilesByExtensions(".xlsx", ".xls");
                result = (from t in files select new StructExelModel { Name = t.Name.Split('.')[0], AbsolutPatch = t.FullName }).ToList();

            }
            catch (Exception ex)
            {
                //throw;
            }

            return result;
        }


        //Создание пакета для всех филлиалов
        public void CreatePackage(string patch)
        {

        }


        //Формирование отчёта для руководства на основании пакета
        public void RootCreateReport()
        {
        }


        //Формирование отчёта для Администрация на основании пакета
        public void AdminCreateReport()
        {
        }


        //Анализатор Даты, по сути нужен при формирования отчётов для Администратора. 
        //Регулирует содание либо нового файла, либо новой книге в старом.
        private void DateAnalisator()
        {

        }

        //Примитивный анализатор (соответствия файлов из пакета с шаблоном)
        //Анализатор имён пакетов
        //Формирование объекта шаблона
        //Формирвание необходимый файлов по шаблону
        //Обработка полученных файлов по шаблону
    }
}
