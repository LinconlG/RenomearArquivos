using System;
using System.IO;

namespace Renomear
{
    class Program
    {
       
        static void Main(string[] args)
        {
            try
            {
                var xlApp = new Microsoft.Office.Interop.Excel.Application();
                var wb = xlApp.Workbooks.Open(@"C:\Users\miojo\Downloads\lista (2).xlsx", ReadOnly: false);
                var ws = wb.Worksheets[1];

                Console.Write("Digite a quatidade de linhas: ");
                int linha = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine();
                Console.Write("Digite a quatidade de colunas: ");
                int coluna = Convert.ToInt32(Console.ReadLine());

                var r = ws.Range["A1"].Resize[linha, coluna];//detectar o tamanho automaticamente ou fazer um jeito de ter input (1)
                var array = r.Value;

                string[] nomesArquivos = new string[linha];
                string[] revisoes = new string[linha];

                for (int i = 1; i <= linha; i++)
                {
                    for (int j = 1; j <= coluna; j++)
                    {
                        string text = Convert.ToString(array[i, j]);

                        if (j == 1)
                        {
                            nomesArquivos[i-1] = text;
                        }
                        else
                        {
                            revisoes[i-1] = text;
                        }
                    }
                }

                xlApp.Quit();
            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }
           

            

        }
    }
}