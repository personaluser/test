namespace Test
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.selectRegFile = new System.Windows.Forms.Button();
            this.regDocLabel = new System.Windows.Forms.TextBox();
            this.btnTest = new System.Windows.Forms.Button();
            this.selectTestFile = new System.Windows.Forms.Button();
            this.testDocLable = new System.Windows.Forms.TextBox();
            this.testRich = new System.Windows.Forms.RichTextBox();
            this.tabaleTitleTreeView = new System.Windows.Forms.TreeView();
            this.tableTreeView1 = new System.Windows.Forms.TreeView();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // selectRegFile
            // 
            this.selectRegFile.Location = new System.Drawing.Point(37, 47);
            this.selectRegFile.Name = "selectRegFile";
            this.selectRegFile.Size = new System.Drawing.Size(90, 35);
            this.selectRegFile.TabIndex = 1;
            this.selectRegFile.Text = "选择规程文件";
            this.selectRegFile.UseVisualStyleBackColor = true;
            this.selectRegFile.Click += new System.EventHandler(this.selectRegFile_Click);
            // 
            // regDocLabel
            // 
            this.regDocLabel.Location = new System.Drawing.Point(159, 55);
            this.regDocLabel.Name = "regDocLabel";
            this.regDocLabel.Size = new System.Drawing.Size(124, 21);
            this.regDocLabel.TabIndex = 9;
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(343, 80);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(183, 54);
            this.btnTest.TabIndex = 14;
            this.btnTest.Text = "运行";
            this.btnTest.UseVisualStyleBackColor = true;
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // selectTestFile
            // 
            this.selectTestFile.Location = new System.Drawing.Point(37, 139);
            this.selectTestFile.Name = "selectTestFile";
            this.selectTestFile.Size = new System.Drawing.Size(90, 35);
            this.selectTestFile.TabIndex = 15;
            this.selectTestFile.Text = "选择目标文件";
            this.selectTestFile.UseVisualStyleBackColor = true;
            this.selectTestFile.Click += new System.EventHandler(this.selectTestFile_Click);
            // 
            // testDocLable
            // 
            this.testDocLable.Location = new System.Drawing.Point(159, 147);
            this.testDocLable.Name = "testDocLable";
            this.testDocLable.Size = new System.Drawing.Size(124, 21);
            this.testDocLable.TabIndex = 16;
            // 
            // testRich
            // 
            this.testRich.Location = new System.Drawing.Point(12, 237);
            this.testRich.Name = "testRich";
            this.testRich.Size = new System.Drawing.Size(126, 269);
            this.testRich.TabIndex = 17;
            this.testRich.Text = "";
            // 
            // tabaleTitleTreeView
            // 
            this.tabaleTitleTreeView.Location = new System.Drawing.Point(500, 47);
            this.tabaleTitleTreeView.Name = "tabaleTitleTreeView";
            this.tabaleTitleTreeView.Size = new System.Drawing.Size(286, 20);
            this.tabaleTitleTreeView.TabIndex = 18;
            // 
            // tableTreeView1
            // 
            this.tableTreeView1.Location = new System.Drawing.Point(497, 12);
            this.tableTreeView1.Name = "tableTreeView1";
            this.tableTreeView1.Size = new System.Drawing.Size(289, 20);
            this.tableTreeView1.TabIndex = 19;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(210, 494);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 12);
            this.label1.TabIndex = 20;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(362, 494);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 12);
            this.label2.TabIndex = 21;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(522, 493);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 12);
            this.label3.TabIndex = 22;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
            this.dataGridView1.BackgroundColor = System.Drawing.Color.White;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(157, 221);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(732, 349);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dataGridView1_CellPainting);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(921, 595);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tableTreeView1);
            this.Controls.Add(this.tabaleTitleTreeView);
            this.Controls.Add(this.testRich);
            this.Controls.Add(this.testDocLable);
            this.Controls.Add(this.selectTestFile);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.regDocLabel);
            this.Controls.Add(this.selectRegFile);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button selectRegFile;
        private System.Windows.Forms.TextBox regDocLabel;
        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.Button selectTestFile;
        private System.Windows.Forms.TextBox testDocLable;
        private System.Windows.Forms.RichTextBox testRich;
        private System.Windows.Forms.TreeView tabaleTitleTreeView;
        private System.Windows.Forms.TreeView tableTreeView1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}

