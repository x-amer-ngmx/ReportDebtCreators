using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportDebtCreators.model;


namespace ReportDebtCreators.enginer
{
    class ExelCreator
    {
        private readonly string _rootPatch;
        public ExelCreator(string rootPatch)
        {
            _rootPatch = rootPatch;
        }

        public List<StructExelModel> ListTemplate()
        {
            List<StructExelModel> result = null;
            try
            {
                var file = new DirectoryInfo($@"{_rootPatch}\Templates\").GetFiles("*");
                result = (from t in file select new StructExelModel {Name = t.Name, AbsolutPatch = t.FullName}).ToList();
            }
            catch (Exception ex)
            {
                // ignored
            }


            return result;
        }

        public List<StructExelModel> ListPackage()
        {
            List<StructExelModel> result = null;

            try
            {
                var file = new DirectoryInfo($@"{_rootPatch}\EnginerFilePackage\").GetDirectories("*");
                result = (from t in file select new StructExelModel { Name = t.Name, AbsolutPatch = t.FullName }).ToList();
            }
            catch (Exception ex)
            {
                // ignored
            }

            return result;
        }
    }
}
