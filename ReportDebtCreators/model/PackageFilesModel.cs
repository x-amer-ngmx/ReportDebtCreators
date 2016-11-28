using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportDebtCreators.model
{
    public class PackageFilesModel
    {
        public StructExelModel pack { get; set; }
        public List<StructExelModel> BrangeFiles { get; set; }
    }
}
