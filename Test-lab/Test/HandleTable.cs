using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application=Microsoft.Office.Interop.Word.Application;
using System.Collections;

namespace Test
{
    class HandleTable
    {
        private Application testWord;
        List<string> regTreeList=new List<string>();
        List<string> targetTreeList = new List<string>();

        ArrayList regMissList = new ArrayList();//记录规程文档缺失表格标题位置
        ArrayList regFaultList = new ArrayList();//记录规程文档出现匹配错误的表格位置
        ArrayList regRightList = new ArrayList();//记录规程文档出现匹配正确的表格位置

        ArrayList targetFaultList = new ArrayList();//记录目标文档出现匹配错误的表格位置
        ArrayList targetRightList = new ArrayList();//记录目标文档出现匹配正确的表格位置
        
        public HandleTable(Application app)
        {
            testWord = app;
        }
        public void contrastTablesOfDocs(Document regulation, Document target, RichTextBox compareBox, TreeView regTableTitleTreeView
                                                  , TreeView targetTableTitleTreeView,Label right,Label error,Label missing)//对规程和目标文档进行表格的对比,对比信息显示在compareBox上
        {
            Boolean hasAttachedTable = false;//是否有附表

            int location = 0;

            Find find;
            Table regTable;
            Table testTable;
            string compareResult = "";//对比结果
            
            string targetTitle = "";//所要查找的目标表格的标题

            string[] regtableOfContents = getTableOfContents(regulation);
            regulation.TablesOfContents[1].Delete();//获取目录之后要删除

            getTargetTreeList(target);
           
            foreach (string s in regtableOfContents)
            {
                if (s.Contains("表"))//如果该标题是离附表最近的标题
                {
                    hasAttachedTable = true;//标记有附表

                    find = getFind(regulation, s);
                    find.Execute();

                    regTreeList.Add(getTitle(s));//逐步添加treeList的元素
                    regTable = getTable(regulation);//获取到规程文档中的表格

                    targetTitle = getNormalTitle(s);//得到目标文档中表的title

                    find = getFind(target, targetTitle);
                    if (find.Execute())
                    {
                        testTable = getTable(target);//获取到目标文档中的表格
                        string result = compareTables(regTable, testTable, targetTitle);
                        if (!result.Equals("正确"))
                        {
                            regFaultList.Add(location);
                            compareResult += result + "\r\r";
                            int targetLocation = targetTitleLocation(targetTitle);
                            targetFaultList.Add(targetLocation);//记录目标文档错误表格标题的位置
                        }
                        else
                        {
                            int targetLocation = targetTitleLocation(targetTitle);
                            targetRightList.Add(targetLocation);
                            regRightList.Add(location);
                        }
                        testWord.Selection.HomeKey(WdUnits.wdStory, Type.Missing);
                    }
                    else
                    {
                        regMissList.Add(location);
                    }
                    location++;
                }
            }
            if (hasAttachedTable == false)
            {
                MessageBox.Show("该文档标题不含附表格式不符");
                return;
            }
            showResult(regTableTitleTreeView, targetTableTitleTreeView, compareBox, compareResult, right, error, missing);   

        }
        private void showResult(TreeView regTableTitleTreeView, TreeView targetTableTitleTreeView, RichTextBox compareBox,
                                                                          string compareResult,Label right, Label error, Label missing)
        {
            compareBox.Text = compareResult;
            setLableText(right, error, missing);

            generateTreeView(regTreeList, regTableTitleTreeView);

            generateTreeView(targetTreeList, targetTableTitleTreeView);

            colorTreeView(targetTableTitleTreeView, regTableTitleTreeView);
        }
        private void setLableText(Label Right, Label Error, Label Missing)
        {
            Right.Text +="匹配正确："+targetRightList.Count;
            Error.Text +="匹配错误："+ regFaultList.Count;
            Missing.Text +="缺失表格："+ regMissList.Count;
        }
        private int targetTitleLocation(string s)
        {
            for (int i = 0; i < targetTreeList.Count; i++)
            {
                if(targetTreeList[i].Contains(s))
                    return i;
            }
            return -1;
        }
        private Find getFind(Document doc, string s)
        {
            doc.Activate();
            testWord.Selection.Find.Text = s;
            return testWord.Selection.Find;
        }
        private string getNormalTitle(string s)
        {
            int start = s.LastIndexOf('×') + 1;
            int end = s.IndexOf('\t');
            return s.Substring(start, end - start);
        }
        private string getTitle(string s)   //获取title,结束符为'\t'
        {
            int end = s.IndexOf('\t');
            return s.Substring(0, end);
        }
        private string[] getTableOfContents(Document doc)
        {
            doc.Activate();
            generateTOC(doc, testWord);//自动生成目录
            Range tocRange = doc.TablesOfContents[1].Range;
            string[] tableOfContents = delimiterTOC(tocRange.Text);
            return tableOfContents;
        }
        private void generateTOC(Document doc, Application app)
        {
            doc.Activate();
            Object oMissing = System.Reflection.Missing.Value;
            Object oTrue = true;
            Object oFalse = false;
            Object oUpperHeadingLevel = "1";
            Object oLowerHeadingLevel = "3";
            Object oTOCTableID = "TableOfContents";

            app.Selection.Start = 0;
            app.Selection.End = 0;//将光标移动到文档开始位置  
            object beginLevel = 2;//目录开始深度  
            object endLevel = 2;//目录结束深度  

            object rightAlignPageNumber = true;// 指定页码右对其  

            doc.TablesOfContents.Add(app.Selection.Range, ref oTrue, ref oUpperHeadingLevel,
                ref oLowerHeadingLevel, ref oMissing, ref oTOCTableID, ref oTrue,
                ref oTrue, ref oMissing, ref oTrue, ref oTrue, ref oTrue);
        }
        private string[] delimiterTOC(string tocText)//生成目录的字符串数组
        {
            string strDelimiters = "\r";
            char[] a_charDelimiter = strDelimiters.ToCharArray();
            string strTOC = tocText.TrimEnd(a_charDelimiter);
            string[] contents = strTOC.Split(a_charDelimiter);
            return contents;
        }
        private Table getTable(Document doc)//移动文档光标到表格并获取表格
        {
            doc.Activate();//使doc文档为活动文档

            Selection currentSelect = testWord.Selection;
            while (currentSelect.Tables.Count == 0)
            {
                currentSelect.GoToNext(WdGoToItem.wdGoToLine);
            }
            Table table = currentSelect.Tables[1];
            return table;
        }
        private void getTargetTreeList(Document doc)//获取目标文档形成TreeView的list
        {
            string s="";
            string[] testTableOfContents = getTableOfContents(doc);
            doc.TablesOfContents[1].Delete();
            for (int i = testTableOfContents.Length-1; i >=0; i--)
            {
                s = testTableOfContents[i];
                if (s.Contains("附表")&&s.Contains("："))
                {
                    targetTreeList.Add(getTitle(s));
                }
                else
                    break;
            }
            targetTreeList.Reverse();
        }
        private string compareTables(Table regTable, Table testTable, string targetTableName)
        {
            int regColumns = regTable.Columns.Count;
            int regRows = regTable.Rows.Count;
            int testColumns = testTable.Columns.Count;
            int testRows = testTable.Rows.Count;
            Cell regCell, testCell;

            String compareResult = "";

            if (regColumns == testColumns && regRows == testRows)
            {
                for (int i = 1; i <= regRows; i++)
                {
                    for (int j = 1; j <= regColumns; j++)
                    {
                        try        //如果规程文档表格当前单元格存在,则目标文档相应表格单元格也应存在且内容相等
                        {
                            regCell = regTable.Cell(i, j);
                            try
                            {
                                testCell = testTable.Cell(i, j);
                                if (!regCell.Range.Text.Equals(//目标文档相应单元格存在，如果单元格内容不相同，则内容不一致
                                   testCell.Range.Text))
                                {
                                    //MessageBox.Show(i+"×"+j+"单元格内容不一致");
                                    compareResult =targetTableName+ i + "×" + j + "单元格内容不一致";
                                    return compareResult;
                                }
                            }
                            catch   //目标文档相应表格单元格不存在
                            {
                                //MessageBox.Show("模式不对1:目标文档相应表格单元格不存在");
                                compareResult = targetTableName + ":目标文档相应表格单元格缺失";
                                return compareResult;
                            }
                        }
                        catch    //如果规程文档表格当前单元格不存在，目标文档相应单元格也须不存在
                        {
                            try
                            {
                                testCell = testTable.Cell(i, j);//如果目标文档相应单元格存在则报错
                                //MessageBox.Show("模式不对2:规程文档目标文档单元格不匹配");
                                compareResult =targetTableName+ ":规程文档目标文档某单元格不匹配";
                                return compareResult;
                            }
                            catch   //目标文档相应表格单元格不存在
                            {
                                continue;
                            }
                        }
                    }
                }
            }
            else
            {
                //MessageBox.Show("模式不对3");
                compareResult =targetTableName+ ":目标文档表格行列数有误";
                return compareResult;
            }
            return "正确";
        }
        private void generateTreeView(List<string>list, TreeView tv)
        {
            if (list.Count == 0)
            {
                tv.Nodes.Add("该文档中未找到附表信息");
            }
            foreach (string s  in list)
            {
                tv.Nodes.Add(s);
            }
        }
        private void colorTreeView(TreeView targetTreeView, TreeView regTreeView)
        {
            foreach(int i in regMissList)
            {
                regTreeView.Nodes[i].BackColor = SysColor.MISS;
            }
            foreach (int i in regFaultList)
            {
                regTreeView.Nodes[i].BackColor = SysColor.ERROR;
            }
            foreach (int i in regRightList)
            {
                regTreeView.Nodes[i].BackColor = SysColor.RIGHT;
            }
            foreach (int i in targetFaultList)
            {
                if (i != -1)
                {
                    targetTreeView.Nodes[i].BackColor = SysColor.ERROR;
                }
            }
            foreach (int i in targetRightList)
            {
                targetTreeView.Nodes[i].BackColor =SysColor.RIGHT;
            }
        }
    }
}
