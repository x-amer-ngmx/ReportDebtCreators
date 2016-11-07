namespace ReportDebtCreators
{
    partial class MainCreatorsForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.TemplateLasts = new System.Windows.Forms.ComboBox();
            this.PackageLasts = new System.Windows.Forms.ComboBox();
            this.ChReportRoot = new System.Windows.Forms.RadioButton();
            this.ChReportAdmin = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.CreatePack = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.PackFromList = new System.Windows.Forms.ComboBox();
            this.PackToList = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.ChPack = new System.Windows.Forms.RadioButton();
            this.ChRangPack = new System.Windows.Forms.RadioButton();
            this.MethodGroup = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.GenirateRepotr = new System.Windows.Forms.Button();
            this.CloseApp = new System.Windows.Forms.Button();
            this.MethodGroup.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // TemplateLasts
            // 
            this.TemplateLasts.FormattingEnabled = true;
            this.TemplateLasts.Location = new System.Drawing.Point(114, 15);
            this.TemplateLasts.MinimumSize = new System.Drawing.Size(121, 0);
            this.TemplateLasts.Name = "TemplateLasts";
            this.TemplateLasts.Size = new System.Drawing.Size(252, 21);
            this.TemplateLasts.TabIndex = 1;
            this.TemplateLasts.SelectedIndexChanged += new System.EventHandler(this.TemplateLasts_SelectedIndexChanged);
            this.TemplateLasts.SelectedValueChanged += new System.EventHandler(this.TemplateLasts_SelectedValueChanged);
            // 
            // PackageLasts
            // 
            this.PackageLasts.FormattingEnabled = true;
            this.PackageLasts.Location = new System.Drawing.Point(6, 36);
            this.PackageLasts.Name = "PackageLasts";
            this.PackageLasts.Size = new System.Drawing.Size(252, 21);
            this.PackageLasts.TabIndex = 5;
            // 
            // ChReportRoot
            // 
            this.ChReportRoot.AutoSize = true;
            this.ChReportRoot.Checked = true;
            this.ChReportRoot.Location = new System.Drawing.Point(170, 6);
            this.ChReportRoot.Name = "ChReportRoot";
            this.ChReportRoot.Size = new System.Drawing.Size(113, 17);
            this.ChReportRoot.TabIndex = 8;
            this.ChReportRoot.TabStop = true;
            this.ChReportRoot.Text = "Для руководства";
            this.ChReportRoot.UseVisualStyleBackColor = true;
            // 
            // ChReportAdmin
            // 
            this.ChReportAdmin.AutoSize = true;
            this.ChReportAdmin.Location = new System.Drawing.Point(170, 34);
            this.ChReportAdmin.Name = "ChReportAdmin";
            this.ChReportAdmin.Size = new System.Drawing.Size(175, 17);
            this.ChReportAdmin.TabIndex = 9;
            this.ChReportAdmin.Text = "Для администратора отчётов";
            this.ChReportAdmin.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Выбор шаблона :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(161, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Выбор пакета для обработки :";
            // 
            // CreatePack
            // 
            this.CreatePack.Location = new System.Drawing.Point(372, 14);
            this.CreatePack.Name = "CreatePack";
            this.CreatePack.Size = new System.Drawing.Size(212, 23);
            this.CreatePack.TabIndex = 2;
            this.CreatePack.Text = "Сформировать пакет по филлиалам";
            this.CreatePack.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 19);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(160, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Метод формирование отчёта :";
            // 
            // PackFromList
            // 
            this.PackFromList.FormattingEnabled = true;
            this.PackFromList.Location = new System.Drawing.Point(30, 15);
            this.PackFromList.Name = "PackFromList";
            this.PackFromList.Size = new System.Drawing.Size(121, 21);
            this.PackFromList.TabIndex = 6;
            this.PackFromList.TabStop = false;
            // 
            // PackToList
            // 
            this.PackToList.FormattingEnabled = true;
            this.PackToList.Location = new System.Drawing.Point(197, 15);
            this.PackToList.Name = "PackToList";
            this.PackToList.Size = new System.Drawing.Size(121, 21);
            this.PackToList.TabIndex = 7;
            this.PackToList.TabStop = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(5, 18);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(19, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "с :";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(166, 18);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(25, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "по :";
            // 
            // ChPack
            // 
            this.ChPack.AutoSize = true;
            this.ChPack.Checked = true;
            this.ChPack.Location = new System.Drawing.Point(170, 6);
            this.ChPack.Name = "ChPack";
            this.ChPack.Size = new System.Drawing.Size(56, 17);
            this.ChPack.TabIndex = 3;
            this.ChPack.TabStop = true;
            this.ChPack.Text = "Пакет";
            this.ChPack.UseVisualStyleBackColor = true;
            // 
            // ChRangPack
            // 
            this.ChRangPack.AutoSize = true;
            this.ChRangPack.Location = new System.Drawing.Point(236, 5);
            this.ChRangPack.Name = "ChRangPack";
            this.ChRangPack.Size = new System.Drawing.Size(76, 17);
            this.ChRangPack.TabIndex = 4;
            this.ChRangPack.Text = "Диапазон";
            this.ChRangPack.UseVisualStyleBackColor = true;
            // 
            // MethodGroup
            // 
            this.MethodGroup.Controls.Add(this.label3);
            this.MethodGroup.Controls.Add(this.ChReportRoot);
            this.MethodGroup.Controls.Add(this.ChReportAdmin);
            this.MethodGroup.Location = new System.Drawing.Point(12, 143);
            this.MethodGroup.Name = "MethodGroup";
            this.MethodGroup.Size = new System.Drawing.Size(372, 57);
            this.MethodGroup.TabIndex = 8;
            this.MethodGroup.TabStop = true;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.PackToList);
            this.panel1.Controls.Add(this.PackFromList);
            this.panel1.Location = new System.Drawing.Point(264, 21);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(337, 47);
            this.panel1.TabIndex = 6;
            this.panel1.TabStop = true;
            this.panel1.Visible = false;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.panel1);
            this.panel2.Controls.Add(this.PackageLasts);
            this.panel2.Controls.Add(this.ChPack);
            this.panel2.Controls.Add(this.ChRangPack);
            this.panel2.Location = new System.Drawing.Point(12, 51);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(503, 79);
            this.panel2.TabIndex = 3;
            this.panel2.TabStop = true;
            // 
            // GenirateRepotr
            // 
            this.GenirateRepotr.Location = new System.Drawing.Point(50, 241);
            this.GenirateRepotr.Name = "GenirateRepotr";
            this.GenirateRepotr.Size = new System.Drawing.Size(188, 23);
            this.GenirateRepotr.TabIndex = 0;
            this.GenirateRepotr.Text = "Запустить формирование отчёта";
            this.GenirateRepotr.UseVisualStyleBackColor = true;
            // 
            // CloseApp
            // 
            this.CloseApp.Location = new System.Drawing.Point(352, 241);
            this.CloseApp.Name = "CloseApp";
            this.CloseApp.Size = new System.Drawing.Size(126, 23);
            this.CloseApp.TabIndex = 10;
            this.CloseApp.Text = "Закрыть приложение";
            this.CloseApp.UseVisualStyleBackColor = true;
            this.CloseApp.Click += new System.EventHandler(this.CloseApp_Click);
            // 
            // MainCreatorsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(595, 281);
            this.Controls.Add(this.CloseApp);
            this.Controls.Add(this.GenirateRepotr);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.MethodGroup);
            this.Controls.Add(this.CreatePack);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TemplateLasts);
            this.Icon = global::ReportDebtCreators.Properties.Resources.favicon;
            this.Name = "MainCreatorsForm";
            this.Text = "Формирование отчётов по должникам.";
            this.TopMost = true;
            this.MethodGroup.ResumeLayout(false);
            this.MethodGroup.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox TemplateLasts;
        private System.Windows.Forms.ComboBox PackageLasts;
        private System.Windows.Forms.RadioButton ChReportRoot;
        private System.Windows.Forms.RadioButton ChReportAdmin;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button CreatePack;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox PackFromList;
        private System.Windows.Forms.ComboBox PackToList;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.RadioButton ChPack;
        private System.Windows.Forms.RadioButton ChRangPack;
        private System.Windows.Forms.Panel MethodGroup;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button GenirateRepotr;
        private System.Windows.Forms.Button CloseApp;
    }
}

