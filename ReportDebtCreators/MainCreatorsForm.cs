using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ReportDebtCreators.enginer;
using ReportDebtCreators.model;

namespace ReportDebtCreators
{
    public partial class MainCreatorsForm : Form
    {
        private IList<StructExelModel> fromPack;
        public MainCreatorsForm(IReadOnlyCollection<StructExelModel> temp, IReadOnlyCollection<StructExelModel> pack)
        {
            InitializeComponent();

            FillCollection(temp, pack);

            PanelUnvis();
        }

        private void TemplateLasts_SelectedValueChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("Changed");
        }

        private void TemplateLasts_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("Changed");
        }

        private void CloseApp_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void ChReportRoot_CheckedChanged(object sender, EventArgs e)
        {
            PanelUnvis();
        }

        private void PanelUnvis()
        {
            var id = 92;

            panel2.Visible = !ChReportRoot.Checked;
            MethodGroup.Location = new Point(MethodGroup.Location.X, ChReportRoot.Checked ? 19 : 111);
            this.Size = new Size(this.Size.Width, ChReportRoot.Checked ? this.Size.Height - id : this.Size.Height + id);
            groupBox2.Size = new Size(groupBox2.Size.Width,
                ChReportRoot.Checked ? groupBox2.Size.Height - id : groupBox2.Size.Height + id);

            GenirateRepotr.Location = new Point(GenirateRepotr.Location.X, ChReportRoot.Checked ? GenirateRepotr.Location.Y - id : GenirateRepotr.Location.Y + id);
            CloseApp.Location = new Point(CloseApp.Location.X, ChReportRoot.Checked ? CloseApp.Location.Y - id : CloseApp.Location.Y + id);
        }

        private void PackageLasts_SelectedIndexChanged(object sender, EventArgs e)
        {
            CountPackFile.Visible = true;

            var select = (StructExelModel)PackageLasts.SelectedItem;

            var res = ExelCreator.ListPackageFiles(select.AbsolutPatch);
            
            CountPackFile.Text = $"Всего в пакете {res.Count} файлов.";
        }

        private void PackFromList_SelectedIndexChanged(object sender, EventArgs e)
        {
            var select = (StructExelModel)PackFromList.SelectedItem;

            var result = fromPack.GetRangePack(
                select.Name.PacNameConvert(),
                fromPack.Max(x => x.DateIndex));

            var cou = result.Sum(exelModel => ExelCreator.ListPackageFiles(exelModel.AbsolutPatch).Count);

            PackToList.DataSource = result;

            if(ChRangPack.Checked)
            CountPackFile.Text = $"Всего в пакете {cou} файлов.";
        }

        private void PackToList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ChRangPack_CheckedChanged(object sender, EventArgs e)
        {
            if (ChRangPack.Checked)
            {
                PackFromList.SelectedIndex = 0;
                PackFromList_SelectedIndexChanged(PackFromList, new EventArgs());

                panelRangePack.Visible = true;
                panelRangePack.Location = new Point(4, 30);
                panelPack.Visible = false;
            }
        }

        private void ChPack_CheckedChanged(object sender, EventArgs e)
        {
            if (ChPack.Checked)
            {
                panelRangePack.Visible = false;
                panelPack.Visible = true;
                PackageLasts.SelectedIndex = 0;
                PackageLasts_SelectedIndexChanged(PackFromList, new EventArgs());
            }
        }

        private void CreatePack_Click(object sender, EventArgs e)
        {
            info.Text = "Формирование пакета для филиалов.";
            var x = (StructExelModel)TemplateLasts.SelectedItem;
            var exl = new ExelEnginer(x.AbsolutPatch, this);
            exl.CreatePackFile();
            info.Text = "";
        }


        private void GenirateRepotr_Click(object sender, EventArgs e)
        {
            info.Text = "Формирование отчёта, ожидайте результат.";
            List<PackageFilesModel> pak = null;
            var x = (StructExelModel)TemplateLasts.SelectedItem;
            var obj = new ExelEnginer(x.AbsolutPatch, this);

            if (ChPack.Checked)
            {
                var selectm = (StructExelModel)PackageLasts.SelectedItem;

                pak = fromPack.GetPack(selectm.Name.PacNameConvert()).GetEngineFList();
            }

            if (ChRangPack.Checked)
            {
                var f = (StructExelModel) PackFromList.SelectedItem;
                var t = (StructExelModel) PackToList.SelectedItem;

                pak = fromPack.GetRangePack(
                    f.Name.PacNameConvert(),
                    t.Name.PacNameConvert()
                ).GetEngineFList();
            }

            obj.CreateReport(pak, ChReportAdmin.Checked);

            info.Text = "";
        }

        public void SetInfoLable(string inf)
        {
            info.Text = inf;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var exll = new ExelCreator();
            var temp = exll.ListTemplate(Program.DirTemp);
            var pack = exll.ListPackage(Program.DirEng);

            FillCollection(temp,pack);
        }

        private void FillCollection(IReadOnlyCollection<StructExelModel> temp, IReadOnlyCollection<StructExelModel> pack)
        {
            fromPack = (IList<StructExelModel>)pack;
           

            TemplateLasts.DataSource = temp;
            TemplateLasts.DisplayMember = "Name";
            TemplateLasts.ValueMember = "AbsolutPatch";

            PackageLasts.DataSource = pack;
            PackageLasts.DisplayMember = "Name";
            PackageLasts.ValueMember = "AbsolutPatch";


            IList<StructExelModel> rr = (from p in fromPack orderby p.DateIndex ascending select p).ToList();

            PackFromList.DataSource = rr;
            PackFromList.DisplayMember = "Name";
            PackFromList.ValueMember = "AbsolutPatch";

            PackToList.DisplayMember = "Name";
            PackToList.ValueMember = "AbsolutPatch";
        }

        /*
private void PackFromList_SelectedIndexChanged_1(object sender, EventArgs e)
{
var select = (StructExelModel)PackageLasts.SelectedItem;

var d = ExelCreator.PackageNameAnalisator(select.Name);

var m = fromPack.Max(x => x.DateIndex);

IList<StructExelModel> result = (from p in fromPack orderby p.DateIndex
    where (p.DateIndex >= d && p.DateIndex <= m)
    select p).ToList();

//IList<StructExelModel> rr = (from p in fromPack orderby p.DateIndex descending select p).ToList();
PackToList.DataSource = result;
PackToList.DisplayMember = "Name";
PackToList.ValueMember = "AbsolutPatch";
}*/
    }
}
