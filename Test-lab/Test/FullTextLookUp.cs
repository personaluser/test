using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
namespace Test
{
    class FullTextLookUp
    {
        public void fullTextCheck(string filePath,Boolean isPathChanged, string keyWord,string oldKeyWord,Boolean isKeyWordChanged, Microsoft.Office.Interop.Word.Application word)
        {
            if (word.Documents.Count == 0)
            {
                word = openDocument(filePath, word);
            }
            else if(isPathChanged)//路径改变则关掉之前文档，打开新的文档
            {
                word.Documents.Close(WdSaveOptions.wdDoNotSaveChanges, Type.Missing, Type.Missing);
                word.Quit();
                word = openDocument(filePath, word);
            }
            Selection currentSelect = word.Selection;
            if (isKeyWordChanged)  //关键字改变的话需要还原文档
            {
                if (oldKeyWord != "")
                {
                    restoreDocument(oldKeyWord, currentSelect);
                }
                replace(keyWord, currentSelect);//替换关键字为高亮显示
            }
            
        }

        private Microsoft.Office.Interop.Word.Application openDocument(string filepath, Microsoft.Office.Interop.Word.Application word)
        {
            Object filename = filepath;
            word.Visible = true;

            object isread = true;
            object isvisible = true;
            object miss = System.Reflection.Missing.Value;
            word.Documents.Open(ref filename, ref miss, ref isread, ref miss, ref miss, ref miss, ref miss, ref miss,
                                              ref miss, ref miss, ref miss, ref isvisible, ref miss, ref miss, ref miss, ref miss);
            return word;
        }
        private void replace(string keyWord, Selection currentSelect)
        {
            executeFindReplace(currentSelect, 12, WdColor.wdColorRed, keyWord);
            currentSelect.Find.Execute();
        }

        private void restoreDocument(string keyWord, Selection currentSelect)
        {
            executeFindReplace(currentSelect, 0, WdColor.wdColorBlack, keyWord);
        }

        private void executeFindReplace(Selection currentSelect, int n, WdColor color, string keyWord)
        {
            object MissingValue = Type.Missing;
            object FindText, ReplaceWith, Replace;//   
            FindText = keyWord;
            ReplaceWith = keyWord;//替换文本  
            Replace = WdReplace.wdReplaceAll;
            currentSelect.Find.Replacement.Font.Color = color;
            currentSelect.Find.Replacement.Font.Bold = n;

            currentSelect.HomeKey(WdUnits.wdStory, Type.Missing);
            currentSelect.Find.Execute(ref FindText, ref MissingValue, ref MissingValue, ref MissingValue,
                                       ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue,
                                       ref ReplaceWith, ref Replace, ref MissingValue, ref MissingValue, ref MissingValue, ref MissingValue);
        }
    }
}
