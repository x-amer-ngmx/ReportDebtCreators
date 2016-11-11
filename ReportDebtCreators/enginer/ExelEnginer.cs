using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Execl = Microsoft.Office.Interop.Excel;

namespace ReportDebtCreators.enginer
{
    

    public class ExelEnginer
    {
        private string pwws;

        public ExelEnginer(string pws)
        {
            pwws = pws;
        }

        //основные методы которые стоит вынести за пределы исполняемого класса

        // Открытие файла,
        // Получение листов
        // Закрытие файла
        // Создание файла
        // Сохранение

        public void CreatePackFile(string template, string pack)
        {
            try
            {
                var kernel = new ExelKernel();
                kernel.OpenFile(template);
                /*var shlist = kernel.GetSheetsList();

                var mmx = shlist[0];

                mmx.Protect("funt");*/

                //kernel.CloseFile();
                kernel.Quit(fname: $@"{ConfigurationManager.AppSettings["rootPachExel"]}\xui.xlsx");

                kernel = null;


            }
            catch (Exception ex)
            {
                
                //throw;
            }


        }

    }
}
