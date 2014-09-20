using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Test
{
    public class TermContainer
    {
        string section;
        private string termName;
        private string termTip;

        public TermContainer(string section,string name, string tip)
        {
            this.section = section;
            this.termName = name;
            this.termTip = tip;
        }

        public string TermName
        {
            get { return termName; }
            set { this.termName = value; }
        }

        public string Section
        {
            get { return section; }
            set { section = value; }
        }

        public string TermTip
        {
            get { return termTip; }
            set { termTip = value; }
        }
    }
}



/*
        public string readPageText(int pageNum, Document doc)
        {
            object objWhat = Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage;
            object objWhich = Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute;
            object objPage = pageNum;

            Range range1 = doc.GoTo(ref objWhat, ref objWhich, ref objPage, Type.Missing);
            Range range2 = range1.GoToNext(Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage);

            object objStart = range1.Start;
            object objEnd = range2.Start;

            string text = doc.Range(ref objStart, ref objEnd).Text;
            return text;
        }*/