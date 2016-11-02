using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ReportDebtCreators.enginer;
using ReportDebtCreators.model;

namespace ReportDebtCreators
{
    public partial class MainCreatorsForm : Form
    {
        public MainCreatorsForm(IReadOnlyCollection<StructExelModel> temp, IReadOnlyCollection<StructExelModel> pack)
        {

            InitializeComponent();

            TemplateLasts.DataSource = temp;
            TemplateLasts.DisplayMember = "Name";
            TemplateLasts.ValueMember = "AbsolutPatch";

            PackageLasts.DataSource = pack;
            PackageLasts.DisplayMember = "Name";
            PackageLasts.ValueMember = "AbsolutPatch";
        }


    }
}
