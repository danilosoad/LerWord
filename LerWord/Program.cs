using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using WordApp;

namespace LerWord
{
    class Program
    {
        public static void Main(string[] args)
        {
            //nome do arquivo
            string fileName = "teste";

            //caminho do arquivo
            string path = string.Format(@"C:\Users\shymm\OneDrive\Área de Trabalho\{0}.docx", fileName);

            var leitura = new MetodosWord();

            var texto = leitura.DocReaderGetAllUntilTitle(path, "teste");
            var texto2 = leitura.DocReaderGetAllButTitle(path, "teste");

            //execução dos métodos

            Console.WriteLine("Pega tudo ate encontrar o proximo titulo igual: ");
            Console.WriteLine(texto);

            Console.WriteLine("Pega tudo menos os titulos :");
            Console.WriteLine(texto2);

            Console.WriteLine("Convertendo arquivo docx para HTMl..");

            leitura.WordToHtml(path, fileName);
            Console.ReadKey();

        }
    }
}