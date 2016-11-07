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

        public List<StructExelModel> ListTemplate(string dir)
        {
            List<StructExelModel> result = null;
            try
            {
                var file = new DirectoryInfo(dir).GetFiles("*.xltx");
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
                result = (from t in file where PackageNameAnalisator(t.Name) select new StructExelModel { Name = t.Name, AbsolutPatch = $"{t.FullName}\\" }).ToList();

                foreach (var model in result)
                {
                    ListPackageFiles(model.AbsolutPatch);
                }
            }
            catch (Exception ex)
            {
                // ignored
            }

            return result;
        }

        public List<StructExelModel> ListPackageFiles(string patch)
        {
            List<StructExelModel> result = null;

            try
            {
                var files = new DirectoryInfo(patch).GetFiles("*.xlsx");
                result = (from t in files select new StructExelModel { Name = t.Name.Split('.')[0], AbsolutPatch = t.FullName }).ToList();

                var ex = new ExelEnginer();
                foreach (var model in result)
                {
                    ex.CreatePackFile(model.AbsolutPatch,"");
                }
            }
            catch (Exception ex)
            {
                //throw;
            }

            return result;
        }

        public void CreatePackage(string patch)
        {

        }

        //
        public void RootCreateReport()
        {
        }

        //
        public void AdminCreateReport()
        {
        }


        //Анализатор Даты, по сути нужен при формирования отчётов для Администратора. 
        //Регулирует содание либо нового файла, либо новой книге в старом.
        private void DateAnalisator()
        {

        }

        //Анализатор имён директорий, 
        private static bool PackageNameAnalisator(string packName)
        {
            DateTime temp;
            var result = DateTime.TryParse(packName, out temp);

            return result;
        }




        //Примитивный анализатор (соответствия файлов из пакета с шаблоном)
        //Анализатор имён пакетов
        //Формирование объекта шаблона
        //Формирвание необходимый файлов по шаблону
        //Обработка полученных файлов по шаблону
    }
}
