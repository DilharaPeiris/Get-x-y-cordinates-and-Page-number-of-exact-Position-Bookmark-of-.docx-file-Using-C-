using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using WordToPDF;


namespace ConsoleApplication1
{
    class Program
    {
        
        static void Main(string[] args)
        {
            

            #region

            System.IO.FileInfo ExecutableFileInfo = new System.IO.FileInfo(System.Reflection.Assembly.GetEntryAssembly().Location);
            object docFileName = System.IO.Path.Combine(ExecutableFileInfo.DirectoryName, "file-sample.doc");


            object nullObject = System.Reflection.Missing.Value;
            Word.Application application = new Word.Application();
            Word.Document document = application.Documents.Open(ref docFileName, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject, ref nullObject);


            string bookmark = "Signature_Test";

            Word.Bookmark bm = document.Bookmarks[bookmark];
            Word.Range range = bm.Range;
            //range.Text = "Hello World";

            float x = range.Information[Word.WdInformation.wdHorizontalPositionRelativeToPage];
            float y = range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];
            int pageNo = range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];

            Console.WriteLine("x  : " + x );
            Console.WriteLine("y  : " + y );
            Console.WriteLine("page No: " + pageNo);

            //doc.Bookmarks.Add(bookmark, range);
            document.Close(ref nullObject, ref nullObject, ref nullObject);
            application.Quit(ref nullObject, ref nullObject, ref nullObject);
            document = null;
            application = null;

            Console.ReadLine();
            #endregion


        }


    }
}
