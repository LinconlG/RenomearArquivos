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


                Console.Write("Insira o diretorio da pasta que contem os arquivos: ");
                string diretorio = Console.ReadLine();
                Console.WriteLine();

                Console.Write("Digite extensão dos arquivos na pasta (com o ponto): ");
                string extensao = Console.ReadLine();
                Console.WriteLine();

                Console.Write("Insira o diretorio do arquivo excel, seguido do nome e extensão: ");
                string diretorioExcel = Console.ReadLine();
                Console.WriteLine();

                Console.Write("Digite a quatidade de linhas: ");
                int linhas = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine();

                int colunas = 2;
                

                DirectoryInfo diretorioPasta = new DirectoryInfo($@"{diretorio}");

                var planilha = new Microsoft.Office.Interop.Excel.Application();
                var wb = planilha.Workbooks.Open($@"{diretorioExcel}", ReadOnly: true);
                var ws = wb.Worksheets[1];
                var r = ws.Range["A1"].Resize[linhas, colunas];
                var array = r.Value;

                string[] nomesArquivos = new string[linhas];
                string[] revisoes = new string[linhas];

                for (int i = 1; i <= linhas; i++) //os dois vetores recebem os nomes e revisoes que estão na planilha
                {
                    for (int j = 1; j <= colunas; j++)
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

                //renomeia os arquivos
                FileInfo[] listaArquivos = diretorioPasta.GetFiles();

                foreach (FileInfo arquivo in listaArquivos)
                {
                    for (int i = 0; i < linhas; i++)
                    {
                        if (arquivo.FullName == $@"{diretorioPasta}\{nomesArquivos[i]}{extensao}")
                        {
                            File.Move(arquivo.FullName, arquivo.FullName.Replace($"{extensao}",$" Rev.{revisoes[i]}{extensao}"));
                            break;
                        }
                    }
                }
                wb.Close();
                planilha.Quit();

                Console.WriteLine("Finalizado!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}