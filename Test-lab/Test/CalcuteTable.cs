using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing;

namespace Test
{
    class CalcuteTable
    {
        private Regex regex = new Regex(@"[A-Z]\r");
        private Regex over = new Regex(@"[A-Z](\d)*");

        private List<string> itemName = new List<string>();

        private List<Decimal> originalValue = new List<Decimal>();
        private List<Decimal> calValue = new List<Decimal>();

        Boolean flag = false;

        public void calTableShowMssing(Table comTable ,Table normalTable,DataGridView dataTableView,RichTextBox richTex)
        {
            int columns = calTable(comTable);
            checkTable(comTable, normalTable,richTex);
            //int i = calculateSum(6, 3, table, 1);
            showTable(columns,dataTableView);
        }
        private int calculateSum(int i, int j, Table table, int col)    //一个主项递归计算（部分递归，效率提高）
        {
            // addValueItem(i, j, table,col);
            int itemRow = i;
            int itemColumn = j;

            i++;
            j++;
            Decimal sum = 0;
            while (true)     //循环计算主项里面的值
            {
                int newLine = 0;

                Decimal value = getNormalValue(i, j + col, table); ;
                sum += value;

                int isMItem = isMainItem(i, j, table);
                if (isMItem == 2)      //遇到子项是主项则计算主项并返回newLine
                {
                    newLine = calculateSum(i, j, table, col);
                    i = newLine;
                }
                else               //如果子项不是主项情况2则行数递增1，继续计算
                {
                    i++;
                }
                if (i > table.Rows.Count)  //如果行数大于表行则终止递归,输出sum
                {
                    System.Diagnostics.Debug.WriteLine(sum + "  dddd");
                    calValue.Add(sum);
                    addValueItem(itemRow, itemColumn, table, col);
                    return i - 1;
                }
                Match m = over.Match(table.Cell(i, j - 1).Range.Text);
                if (m.Length == 0)//如果当前主项计算完毕，打印计算值，并返回新的计算行(左边单元格匹配不到大写字母）
                {
                    System.Diagnostics.Debug.WriteLine(sum + "  ssss");
                    calValue.Add(sum);
                    addValueItem(itemRow, itemColumn, table, col);

                    if (i == table.Rows.Count)//遍历到最低端
                    {
                        flag = true;//结束完一轮计算
                        return table.Rows.Count;
                    }
                    return i;
                }
            }
        }
        private int getItems(int i, int j, Table table, List<string> itemNames)    //一个主项递归计算（部分递归，效率提高）
        {
            addMItems(i, j, table, itemNames);
            i++;
            j++;

            while (true)     //循环计算主项里面的值
            {
                int newLine = 0;

                addSubItems(i, j, table, itemNames);
                int isMItem = isMainItem(i, j, table);
                if (isMItem == 2)      //遇到子项是主项则计算主项并返回newLine
                {
                    newLine = getItems(i, j, table, itemNames);
                    i = newLine;
                }
                else               //如果子项不是主项情况2则行数递增1，继续计算
                {
                    i++;
                }
                if (i > table.Rows.Count)  //如果行数大于表行则终止递归,输出sum
                {
                    return i - 1;
                }
                Match m = over.Match(table.Cell(i, j - 1).Range.Text);
                if (m.Length == 0)//如果当前主项计算完毕，打印计算值，并返回新的计算行(左边单元格匹配不到大写字母）
                {

                    if (i == table.Rows.Count)//遍历到最低端
                    {
                        return table.Rows.Count;
                    }
                    return i;
                }
            }
        }
        private int isMainItem(int i, int j, Table table)
        {
            try
            {
                string cell = table.Cell(i + 1, j - 1).Range.Text;
                cell = getNormalCell(cell);
                if (cell.Equals("其中"))
                {
                    return 2;
                }
                cell = table.Cell(i, j - 1).Range.Text;
                Match m = regex.Match(cell);
                if (m.Length != 0)//如果左边是单个大写字母,则为主项情况1
                {
                    return 1;
                }

            }
            catch   //表格单元格不存在
            {

            }
            return 0;
        }
        private string getNormalCell(string cell)
        {
            int end = cell.IndexOf('\r');
            return cell.Substring(0, end);
        }
        private Decimal getNormalValue(int i, int j, Table table)
        {
            Decimal value = 0;
            string cell = table.Cell(i, j).Range.Text;
            cell = getNormalCell(cell);
            if (cell.Contains("-"))
            {
                return value;
            }
            if (cell.Contains("%"))
            {
                int end = cell.IndexOf('%');
                cell = cell.Substring(0, end);
            }
            value = Convert.ToDecimal(cell);
            return value;
        }

        private void addValueItem(int i, int j, Table table, int col)
        {
            if (flag == false)
            {
                string item = table.Cell(i, j).Range.Text;
                item = getNormalCell(item);
                itemName.Add(item);
            }
            Decimal value = getNormalValue(i, j + col, table);
            originalValue.Add(value);
        }
        private void addMItems(int i, int j, Table table, List<string> itemNames)
        {
            string item = table.Cell(i, j).Range.Text;
            item = getNormalCell(item);
            if (itemNames.Count == 0)
            {
                itemNames.Add(item);
            }
            else if (!itemNames[itemNames.Count - 1].Equals(item))
            {
                itemNames.Add(item);
            }
        }
        private void addSubItems(int i, int j, Table table, List<string> itemNames)
        {
            string item = table.Cell(i, j).Range.Text;
            item = getNormalCell(item);
            itemNames.Add(item);
        }
        private int[] getStartLocation(Table table)
        {
            int[] location = new int[2];
            int rows = table.Rows.Count;
            int columns = table.Columns.Count;

            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= columns; j++)
                {
                    try
                    {
                        string cell = table.Cell(i, j).Range.Text;
                        Match m = regex.Match(cell);
                        if (m.Length != 0)//匹配到第一个大写字母则终止查询
                        {
                            location[0] = i;
                            location[1] = j + 1;
                            return location;
                        }
                    }
                    catch
                    {

                    }
                }
            }
            return location;
        }
        private int calColumn(Table table, int i, int j)//计算有几列值需要被计算
        {
            int columns = 0;
            while (true)
            {
                try
                {
                    j++;
                    Cell cell = table.Cell(i, j);
                    columns++;
                }
                catch
                {
                    return columns;
                }
            }
        }


        private int calTable(Table table) //返回计算列数
        {
            int[] startLocation = getStartLocation(table);
            int rows = table.Rows.Count;
            int columns = calColumn(table, startLocation[0], startLocation[1]);
            for (int i = 1; i <= columns; i++)    //有columns列值需要计算
            {
                int line = startLocation[0];
                while (true)
                {
                    if (isMainItem(line, startLocation[1], table) == 1)   //如果当前单元格是主项情况1则继续寻找知道遇到主项2,再细算
                    {
                        line = findItem2(table, line, startLocation[1]);
                        if (line == 0)
                        {
                            break;
                        }
                    }
                    line = calculateSum(line, startLocation[1], table, i);
                    if (line == rows)
                    {
                        break;
                    }
                }
            }
            return columns;
        }
        private void checkTable(Table comTable, Table normalTable,RichTextBox missingView)
        {
            List<string> list = new List<string>();
            List<string> list1 = new List<string>();

            for (int i = 1; i <= 2; i++)
            {
                Table table;
                if (i == 1)
                {
                    table = normalTable;
                }
                else
                {
                    table = comTable;
                }
                int[] startLocation = getStartLocation(table);
                int rows = table.Rows.Count;

                int line = startLocation[0];

                while (true)
                {
                    if (isMainItem(line, startLocation[1], table) == 1)   //如果当前单元格是主项情况1则继续寻找知道遇到主项2,再细算
                    {
                        if (i == 1)
                        {
                            addMItems(line, startLocation[1], table, list);
                        }
                        else
                        {
                            addMItems(line, startLocation[1], table, list1);
                        }

                        line = findItem2(table, line, startLocation[1]);
                        if (line == 0)
                        {
                            break;
                        }
                    }

                    if (i == 1)
                    {
                        line = getItems(line, startLocation[1], table, list);
                    }
                    else
                    {
                        line = getItems(line, startLocation[1], table, list1);
                    }
                    if (line == rows)
                    {
                        break;
                    }
                }
            }
            List<string> missingList = checkMissing(list, list1);
            if (missingList.Count == 0)
            {
                missingView.Text = "无缺失条目";
            }
            else
            {
                string s = "";
                for (int i = 0; i < missingList.Count; i++)
                {
                    s = missingList[i] + "\n";
                }
                missingView.Text = s;
            }

        }

        private List<string> checkMissing(List<string> list, List<string> list1)
        {
            List<string> missingList = new List<string>();
            for (int i = 0; i < list.Count; i++)
            {
                if (!list1.Contains(list[i]))
                {
                    missingList.Add(list[i]);
                }
            }
            return missingList;
        }
        private int findItem2(Table table, int line, int column)
        {
            while (true)//循环直到找到主项情况2
            {
                line++;
                if (isMainItem(line, column, table) == 2)
                {
                    return line;
                }
                if (line == table.Rows.Count)   //大于表行还未找到则终止
                {
                    flag = true;//结束完一轮计算
                    return 0;
                }
            }
        }

        private void showTable(int columns,DataGridView dataTableView)
        {
            for (int i = 0; i < 2 * columns + 1; i++)
            {
                dataTableView.Columns.Add(new DataGridViewTextBoxColumn());
            }
            int rows = itemName.Count;
            for (int i = 0; i < rows + 1; i++)
            {
                int index = dataTableView.Rows.Add();
                if (i == 0)
                {
                    for (int j = 1; j <= columns; j++)
                    {
                        int start = (j - 1) * 2;
                        dataTableView.Rows[index].Cells[start + 1].Value = j + "(原)";
                        dataTableView.Rows[index].Cells[start + 2].Value = j + "(计)";
                    }
                }
                else
                {
                    dataTableView.Rows[index].Cells[0].Value = itemName[i - 1];
                }
            }
            for (int j = 1; j <= columns; j++)
            {
                int start = (j - 1) * rows;
                for (int i = 1; i <= rows; i++)
                {
                    dataTableView.Rows[i].Cells[j * 2 - 1].Value = originalValue[start + i - 1];
                    dataTableView.Rows[i].Cells[j * 2].Value = calValue[start + i - 1];
                    if (originalValue[start + i - 1] != calValue[start + i - 1])
                    {
                        dataTableView.Rows[i].Cells[j * 2].Style.BackColor = Color.Red;
                    }
                }
            }
        }
        public Table getTableByName(Document doc, string title, Microsoft.Office.Interop.Word.Application testWord)  //testWord下的doc文档下
        //找到标题为title的表格
        {
            doc.Activate();
            testWord.Selection.HomeKey(WdUnits.wdStory, Type.Missing);
            Find find = testWord.Selection.Find;
            find.Text = title;
            find.Execute();

            Selection currentSelect = testWord.Selection;
            if (currentSelect.Range.Text == null)
            {
                return null;
            }
            while (currentSelect.Tables.Count == 0)
            {
                currentSelect.GoToNext(WdGoToItem.wdGoToLine);
            }
            Table table = currentSelect.Tables[1];
            return table;
        }
        /*完全递归版
private int calTable(Table table) //返回计算列数
{
    int[] startLocation = getStartLocation(table);
    int columns = calColumn(table, startLocation[0], startLocation[1]);
    for (int i = 1; i <= columns; i++)
    {
       calculateSum(startLocation[0], startLocation[1], table, i);
    }
    return columns;
}
 * */
        /*完全递归计算表格
        private int calculateSum(int i,int j,Table table,int col)
        {
            int itemRow = i;
            int itemColumn = j;

            i++;
            j++;
            Decimal sum = 0;
            while (true)     //循环计算主项里面的值
            {
                int newLine=0;
                
                Decimal value=getNormalValue(i,j+col,table);;
                sum += value;
                
                int isMItem = isMainItem(i, j, table);
                if (isMItem==2)      //遇到子项是主项则计算主项并返回newLine
                {
                    newLine = calculateSum(i, j, table,col);
                    i = newLine;
                }
                else               //如果子项不是主项情况2则行数递增1，继续计算
                {
                    i++;
                }
                if (i > table.Rows.Count)  //如果行数大于表行则终止递归,输出sum
                {
                    System.Diagnostics.Debug.WriteLine(sum+"  dddd");
                    calValue.Add(sum);
                    addValueItem(itemRow, itemColumn, table, col);
                    return i-1;
                }
                Match m = over.Match(table.Cell(i, j - 1).Range.Text);
                if (m.Length==0)//如果当前主项计算完毕，打印计算值，并返回新的计算行(左边单元格匹配不到大写字母）
                {
                    System.Diagnostics.Debug.WriteLine(sum+"  ssss");
                    calValue.Add(sum);
                    addValueItem(itemRow, itemColumn, table, col);
                    
                    if (i == table.Rows.Count)//遍历到最低端
                    {
                        flag = true;//结束完一轮计算
                        return 0;
                    }

                    if (isMainItem(i, j - 1, table) == 2)
                    {
                        calculateSum(i, j - 1, table,col);//   如果（i，j-1）是主项2继续计算下一个主项
                    }
                    else if (isMainItem(i, j-1, table) == 1)   //如果当前单元格是主项情况1则继续寻找知道遇到主项2,再细算
                    {
                        while (true)//循环直到找到主项情况2
                        {
                            i++;
                            if (isMainItem(i, j - 1, table) == 2)
                            {
                                calculateSum(i, j - 1, table,col);//   如果（i，j-1）是主项2继续计算下一个主项
                                return 0;
                            }
                            if (i > table.Rows.Count)   //大于表行还未找到则终止
                            {
                                flag = true;//结束完一轮计算
                                return 0;
                            }
                        }
                    }
                    return i;
                }  
            }
        }*/

    }
}
