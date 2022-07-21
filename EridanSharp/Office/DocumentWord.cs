using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace EridanSharp
{
    namespace Office
    {
        public class DocumentWord
        {
            private Word.Application word;
            private Word.Document document;
            private object missing;
            private DocumentProperties documentProperties;
            public DocumentWord()
            {
                //Create an instance for word app.
                word = new Word.Application();

                //Set animation status for word application.
                word.ShowAnimation = false;

                //Set status for word application is to be visible or not.  
                word.Visible = false;

                //Create a missing variable for missing value.
                missing = System.Reflection.Missing.Value;

                //Create a new document.
                document = word.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            }
            public DocumentWord(string pathToFile)
            {
                //Create an instance for word app.
                word = new Word.Application();

                //Set animation status for word application.
                word.ShowAnimation = false;

                //Set status for word application is to be visible or not.  
                word.Visible = false;

                //Create a missing variable for missing value.
                missing = System.Reflection.Missing.Value;

                document = word.Documents.Open(pathToFile);
            }

            public DocumentWord(DocumentProperties documentProperties) : this()
            {
                this.documentProperties = documentProperties;
            }

            public void Replace(object findText, object replaceWithText)
            {
                //options
                object matchCase = false;
                object matchWholeWord = true;
                object matchWildCards = false;
                object matchSoundsLike = false;
                object matchAllWordForms = false;
                object forward = true;
                object format = false;
                object matchKashida = false;
                object matchDiacritics = false;
                object matchAlefHamza = false;
                object matchControl = false;
                object read_only = false;
                object visible = true;
                object replace = 2;
                object wrap = 1;
                //execute find and replace
                word.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                    ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                    ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            }

            public void SaveToFile()
            {
                if(documentProperties != null)
                {
                    //Save to file with properties.
                document.SaveAs(
                    documentProperties.FileName,
                    documentProperties.FileFormat,
                    documentProperties.LockComments,
                    documentProperties.Password,
                    documentProperties.AddToRecentFiles,
                    documentProperties.Encoding,
                    documentProperties.InsertLineBreaks,
                    documentProperties.AllowSubstitutions,
                    documentProperties.LineEnding,
                    documentProperties.AddBiDiMarks
                    );
                }
                else
                {
                    document.SaveAs("file.docx");
                }
            }

            public void SaveToFile(string fileName)
            {
                //Save to file without properties.
                document.SaveAs(fileName);
            }
        }
    }
}
