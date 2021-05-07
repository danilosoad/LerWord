using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Word = Microsoft.Office.Interop.Word;

// @TODO: trazer array de titulos, selecionar um titulo e trazer o texto abaixo

namespace LerWord
{
    public class MetodosWord
    {
        //pega todas as linhas menos as com titulo informado
        public string DocReaderGetAllButTitle(string fileLocation, string headingText)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object miss = System.Reflection.Missing.Value;
            object path = fileLocation;
            object readOnly = true;
            Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            string totaltext = "";
            int paraCount = docs.Paragraphs.Count;

            for (int i = 1; i <= paraCount; i++)
            {

                if (docs.Paragraphs[i].Range.Text.ToString().TrimEnd('\r').ToUpper() != headingText.ToUpper())
                {
                    totaltext += " \r\n " + HttpUtility.HtmlDecode(docs.Paragraphs[i].Range.Text.ToString());

                }
            }

            if (totaltext == "")
            {
                totaltext = "No such data found!";
            }

            docs.Close();
            word.Quit();

            return totaltext;
        }

        //Pega todas as linhas até encontrar outro titulo igual
        public string DocReaderGetAllUntilTitle(string fileLocation, string headingText)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object miss = System.Reflection.Missing.Value;
            object path = fileLocation;
            object readOnly = true;
            Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            string totaltext = "";
            int ind = 0;
            int paraCount = docs.Paragraphs.Count;

            for (int i = 1; i < paraCount; i++)
            {
                if (docs.Paragraphs[i].Range.Text.ToString().TrimEnd('\r').ToUpper() != headingText.ToUpper())
                {
                    totaltext += " \r\n " + docs.Paragraphs[i].Range.Text.ToString();
                }
                else if (docs.Paragraphs[i].Range.Text.ToString().TrimEnd('\r').ToUpper() == headingText.ToUpper() && ind > 0)
                {
                    break;
                }
                else
                {
                    ind++;
                }

            }

            if (totaltext == "")
            {
                totaltext = "No such data found!";
            }

            docs.Close();
            word.Quit();

            return totaltext;
        }

        public void WordToHtml(string path, string fileName)
        {
            object TargetPath = string.Format(@"C:\Users\shymm\OneDrive\Área de Trabalho\{0}",fileName);
            object miss = System.Reflection.Missing.Value;

            object readOnly = true;

            Application word = new Application();
            Document document = new Document();

            document = word.Documents.Open(path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            object format = Word.WdSaveFormat.wdFormatHTML;

            word.ActiveDocument.SaveAs(ref TargetPath, ref format,
                    ref miss, ref miss, ref miss,
                    ref miss, ref miss, ref miss,
                    ref miss, ref miss, ref miss,
                    ref miss, ref miss, ref miss,
                    ref miss, ref miss);

            word.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            word.Quit();
        }

    }
}
