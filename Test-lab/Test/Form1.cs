using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.IO;
using System.Text.RegularExpressions;

namespace Test
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Word.Application testWord = new Microsoft.Office.Interop.Word.Application();
        private Document testDoc = new Document();
        private Document regDoc = new Document();
        private HandleDocument handleDocument = new HandleDocument();

        private Regex regex = new Regex(@"[A-Z]\r");
        private Regex over = new Regex(@"[A-Z](\d)*");

        private List<string> itemName = new List<string>();

        private List<Decimal> originalValue = new List<Decimal>();
        private List<Decimal> calValue = new List<Decimal>();
        private List<string> compareItemName = new List<string>();   //目标表格中子项名字
        private List<string> normalSubitemName = new List<string>();//标准表格中子项名字

        Boolean flag = false;

        public Form1()
        {
            InitializeComponent();
        }


        private string regFileName = "";//规程文档选择
        private void selectRegFile_Click(object sender, EventArgs e)
        {
            string temp = regFileName;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word文件|*.doc";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                regFileName = openFileDialog.FileName;
                if (regFileName != temp && temp != "")
                {
                    regDocLabel.Text = regFileName;
                }
                else
                {
                    regDocLabel.Text = regFileName;

                }
            }
        }
        private string testFileName = "";//测试文档选择
        private void selectTestFile_Click(object sender, EventArgs e)
        {
            string temp = testFileName;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word文件|*.doc";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                testFileName = openFileDialog.FileName;
                if (regFileName != temp && temp != "")
                {
                    testDocLable.Text = testFileName;
                }
                else
                {
                    testDocLable.Text = testFileName;

                }
            }
        }

        //Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
        private void btnTest_Click(object sender, EventArgs e)
        {
            buildShowTable1(dataGridView1);
        }
        private void quit()
        {
            //word.Quit();
            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName("WINWORD"))
            {
                p.Kill();
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            testWord.Quit();
        }
        private void buildShowTable1(DataGridView dataTableView)
        {
            for (int i = 0; i < 2 * 9 + 1; i++)
            {
                dataTableView.Columns.Add(new DataGridViewTextBoxColumn());
            }
            dataTableView.Rows.Add();
            for (int j = 1; j <= 3; j++)
            {
                string name = "";
                if (j == 1)
                {
                    name = "面积（公顷）";
                }
                else if (j == 2)
                {
                    name = "占城市建设用地比重";
                }
                else
                {
                    name = "人均（m2/人）";
                }
                int start = (j - 1) * 6;
                dataTableView.Rows[0].Cells[start + 1].Value = name + "现状";
                dataTableView.Rows[0].Cells[start + 2].Value = name + "现状";
                dataTableView.Rows[0].Cells[start + 3].Value = name + "近期";
                dataTableView.Rows[0].Cells[start + 4].Value = name + "近期";
                dataTableView.Rows[0].Cells[start + 5].Value = name + "远期";
                dataTableView.Rows[0].Cells[start + 6].Value = name + "远期";
            }
            dataTableView.Rows.Add();
            for (int i = 1; i <= 9; i++)
            {
                int start = (i - 1) * 2;
                dataTableView.Rows[1].Cells[start + 1].Value = "原始数据";
                dataTableView.Rows[1].Cells[start + 2].Value = "校验结果";
            }
            dataTableView.Rows[0].Cells[0].Value = "城市建设用地平衡表";
            dataTableView.Rows[1].Cells[0].Value = "城市建设用地平衡表";
            dataTableView.Rows[2].Cells[0].Value = "城市建设用地平";
        }
        private void buildShowTable(DataGridView dataTableView)
        {
            for (int i = 0; i <= 3; i++)
            {
                dataTableView.Columns.Add(new DataGridViewTextBoxColumn());
            }


            dataTableView.Rows.Add();

            dataTableView.Rows[0].Cells[0].Value = 1;
            dataTableView.Rows[0].Cells[1].Value = 1;
            dataTableView.Rows[0].Cells[2].Value = 2;
            dataTableView.Rows[0].Cells[3].Value = 2;
            dataTableView.Rows[1].Cells[0].Value = 1;
            dataTableView.Rows[1].Cells[1].Value = 1;
            dataTableView.Rows[1].Cells[2].Value = 2;
            dataTableView.Rows[1].Cells[3].Value = 2;


        }
         
        Boolean f = false;
        int start = 1;
        /*
        #region 统计行单元格绘制
       
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex ==0)
            {
                e.CellStyle.Font = new System.Drawing.Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Bold);
                e.CellStyle.WrapMode = DataGridViewTriState.True;
                DGVOper.MerageRowSpan(dataGridView1, e, 1, 2);
                DGVOper.MerageRowSpan(dataGridView1, e, 3, 4);
                DGVOper.MerageRowSpan(dataGridView1, e, 5, 6);
                DGVOper.MerageRowSpan(dataGridView1, e, 7, 8);
                DGVOper.MerageRowSpan(dataGridView1, e, 9, 10);
                DGVOper.MerageRowSpan(dataGridView1, e, 11, 12);
                DGVOper.MerageRowSpan(dataGridView1, e, 13, 14);
                DGVOper.MerageRowSpan(dataGridView1, e, 15, 16);
                DGVOper.MerageRowSpan(dataGridView1, e, 17, 18);
                if (start <= e.ColumnIndex && e.ColumnIndex <= start+1)
                {
                    f = DGVOper.MerageRowSpan(dataGridView1, e, start, start+1);
                    if (f == true)
                    {
                        DGVOper.clear();
                        start+=2;
                      
                    }
                }
            }
        }
        #endregion
        */
        
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex != -1 && e.RowIndex == 0)
            {
                using
                    (
                    Brush gridBrush = new SolidBrush(this.dataGridView1.GridColor),
                    backColorBrush = new SolidBrush(e.CellStyle.BackColor)
                    )
                {
                    using (Pen gridLinePen = new Pen(gridBrush))
                    {
                        try
                        {
                            // 清除单元格
                            e.CellStyle.Font = new System.Drawing.Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Regular);
                            e.CellStyle.WrapMode = DataGridViewTriState.True;
                            e.Graphics.FillRectangle(backColorBrush, e.CellBounds);

                            e.Handled = true;

                            // 画 Grid 边线（仅画单元格的底边线和右边线）
                            //   如果下一列和当前列的数据不同，则在当前的单元格画一条右边线
                            if (e.ColumnIndex < dataGridView1.Columns.Count - 1 &&
                            dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex + 1].Value.ToString() !=
                            e.Value.ToString())
                                e.Graphics.DrawLine(gridLinePen, e.CellBounds.Right - 1,
                                e.CellBounds.Top - 1, e.CellBounds.Right - 1,
                                e.CellBounds.Bottom - 1);
                            //画最后一条记录的右边线 
                            if (e.ColumnIndex == dataGridView1.Columns.Count - 1)
                                e.Graphics.DrawLine(gridLinePen, e.CellBounds.Right - 1, e.CellBounds.Top - 1, e.CellBounds.Right - 1, e.CellBounds.Bottom - 1);
                            //画底边线
                            e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left - 1,
                            e.CellBounds.Bottom - 1, e.CellBounds.Right - 1,
                            e.CellBounds.Bottom - 1);

                            // 画（填写）单元格内容，相同的内容的单元格只填写第一个
                            if (e.Value != null)
                            {
                                if (e.ColumnIndex == dataGridView1.Columns.Count - 1
                                     && dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString() !=e.Value.ToString())
                                {
                                    e.Graphics.DrawString(e.Value.ToString(), e.CellStyle.Font,
                                        Brushes.Black, e.CellBounds.Left ,
                                        e.CellBounds.Top+4 , StringFormat.GenericDefault);
                                }
                                if (e.ColumnIndex > 0 &&
                                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex +1].Value.ToString() ==
                                e.Value.ToString())
                                {
                                    e.Graphics.DrawString(e.Value.ToString(), e.CellStyle.Font,
                                          Brushes.Black, e.CellBounds.Left ,
                                          e.CellBounds.Top+4, StringFormat.GenericDefault);
                                }
                                if (e.ColumnIndex > 0 &&
                                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex + 1].Value.ToString() !=
                                e.Value.ToString() && dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString() !=
                                e.Value.ToString())
                                {
                                    e.Graphics.DrawString(e.Value.ToString(), e.CellStyle.Font,
                                          Brushes.Black, e.CellBounds.Left,
                                          e.CellBounds.Top+4, StringFormat.GenericDefault);
                                }

                            }
                            //e.Handled = true;
                        }
                        catch
                        {
                        }
                    }
                }


            }
            if (e.ColumnIndex == 0 && e.RowIndex >= 0)
            {
                using (
                    Brush gridBrush = new SolidBrush(this.dataGridView1.GridColor),
                    backColorBrush = new SolidBrush(e.CellStyle.BackColor))
                {
                    using (Pen gridLinePen = new Pen(gridBrush))
                    {
                        // 擦除原单元格背景
                        e.Graphics.FillRectangle(backColorBrush, e.CellBounds);
                        ////绘制线条,这些线条是单元格相互间隔的区分线条,
                        ////因为我们只对列name做处理,所以datagridview自己会处理左侧和上边缘的线条
                        if (e.RowIndex != 0)
                        {
                            if (e.Value.ToString() != this.dataGridView1.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.ToString())
                            {
                                e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left, e.CellBounds.Top - 1,
                                e.CellBounds.Right - 1, e.CellBounds.Top - 1);//上边缘的线
                                //绘制值
                                if (e.Value != null)
                                {
                                    e.Graphics.DrawString((String)e.Value, e.CellStyle.Font,
                                        Brushes.Black, e.CellBounds.Left,
                                        e.CellBounds.Top+4, StringFormat.GenericDefault);
                                }
                            }
                        }
                        else
                        {
                            //e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left, e.CellBounds.Bottom - 1,
                                //e.CellBounds.Right - 1, e.CellBounds.Bottom - 1);//下边缘的线
                            //绘制值
                            if (e.Value != null)
                            {
                                e.Graphics.DrawString((String)e.Value, e.CellStyle.Font,
                                    Brushes.Black, e.CellBounds.Left,
                                    e.CellBounds.Top+4, StringFormat.GenericDefault);
                            }
                        }
                        if (e.RowIndex == dataGridView1.Rows.Count - 1)
                        {
                            e.Graphics.DrawLine(gridLinePen, e.CellBounds.Left, e.CellBounds.Bottom - 1,
                            e.CellBounds.Right - 1, e.CellBounds.Bottom - 1);//下边缘的线
                        }
                        //右侧的线
                        e.Graphics.DrawLine(gridLinePen, e.CellBounds.Right - 1,
                            e.CellBounds.Top, e.CellBounds.Right - 1,
                            e.CellBounds.Bottom - 1);
                        e.Handled = true;
                    }
                }
            }
        }
    }
}
