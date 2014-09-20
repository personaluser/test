using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Application =Microsoft.Office.Interop.Word.Application;
using System.Windows.Forms;
namespace Test
{
    class HandleDocument
    {
        static object unknow = Type.Missing;
        static Object readOnly = true;
        static Object isVisibale = false;
        static  Object saveChanges = false;

        public Document openDocument(string filepath,Application word)
        {
            Object filename = filepath;

            return word.Documents.Open(ref filename, ref unknow,
                        ref unknow, ref unknow, ref unknow, ref unknow,
                       ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                       ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);
        }
          public Application openDocument(string filepath)
        {
            Object filename = filepath;
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true;

            object isread = true;
            object isvisible = true;
            object miss = System.Reflection.Missing.Value;
            word.Documents.Open(ref filename, ref miss, ref isread, ref miss, ref miss, ref miss, ref miss, ref miss,
                                              ref miss, ref miss, ref miss, ref isvisible, ref miss, ref miss, ref miss, ref miss);
            return word;
        }
        public Dictionary<string, List<TermContainer>> getTestDocTOC(Document doc,Application app)
        {
            generateTOC(doc,app);
            Range tocRange = doc.TablesOfContents[1].Range;
            string[] tocContents = delimiterTOC(tocRange.Text);
           
            Dictionary<string, List<TermContainer>> dict = new Dictionary<string, List<TermContainer>>();
            string chapter = "";
            string section = "";
            
            foreach (string s in tocContents)
            {
                if (isLevelChapter(s))
                {
                     chapter = s;
                     dict.Add(s, new List<TermContainer>());
                }
                else
                {
                     if (isLevelSection(s))
                     {
                         section = s;
                     }
                     else
                     {
                         TermContainer term = new TermContainer(section, s, null);
                         dict[chapter].Add(term);
                     }
                }
                
            }
            return dict;
        }

        private string[] delimiterTOC(string tocText)
        {
            string strDelimiters = "\r";
            char[] a_charDelimiter = strDelimiters.ToCharArray();
            string strTOC = tocText.TrimEnd(a_charDelimiter);
            string[] contents = strTOC.Split(a_charDelimiter);
            return contents;
        }

        /*
        * if row is Level 1 ==> 判断显示章的行
        */
        private bool isLevelChapter(string s)
        {
            if (s.Contains("第") && s.Contains("章"))
                return true;
            return false;
        }

        /*
         * if row is Level section ==> 判断显示节的行
         */
        private bool isLevelSection(string s)
        {
            if (s.Contains("节") && s.Contains("第"))
                return true;
            return false;
        }

        public void generateTOC(Document doc, Application app)
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
        public void quit(Application word, Document doc)
        {
            Object saveChanges = false;
            object unknow = Type.Missing;

            doc.Close(ref saveChanges, ref unknow, ref unknow);
            word.Quit(ref saveChanges, ref unknow, ref unknow);
        }
/*        internal Document openDocument(string testFileName, Microsoft.Office.Interop.Word.Application testWord)
        {
            throw new NotImplementedException();
        }*/
    }
}
