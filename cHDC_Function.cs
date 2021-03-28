using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HDC1
{
    class cHDC_Function
    {
        /// <summary>
        /// apply template từ file template vào file hiện hành
        /// </summary>
        public static void ApplyTemplate(string pathFileTemplate)
        {

            Document thisdoc = Globals.ThisAddIn.Application.ActiveDocument;
            //thisdoc.Activate();
            string sName = thisdoc.FullName;
            Range rng = thisdoc.Range();

            rng.Copy();
            object missing = Type.Missing;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document thatdoc = app.Documents.Open("@" + pathFileTemplate);
            Range rng1 = thatdoc.Range();
            rng1.InsertFile(sName);
            //thatdoc.Content.Paste();
        }
    }
}
