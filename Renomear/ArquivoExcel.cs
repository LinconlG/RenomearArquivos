using System;
using System.IO;

namespace Renomear
{
    class ArquivoExcel
    {
        public static DirectoryInfo diretorioPasta { get; private set; }
        public static string diretorio { get; private set; }
        public static string diretorioPlanilha { get; private set; }
        public static string extensao { get; private set; }
        public static int linhas { get; private set; }
        public static int colunas { get; private set; }

        public ArquivoExcel(string diretorio, string extensao, string diretorioPlanilha, int linhas, int colunas)
        {
            diretorioPasta = new DirectoryInfo($@"{diretorio}");

            var planilha = new Microsoft.Office.Interop.Excel.Application();
            var wb = planilha.Workbooks.Open($@"{diretorioPlanilha}", ReadOnly: true);
            var ws = wb.Worksheets[1];
            var r = ws.Range["A1"].Resize[linhas, colunas];
            var array = r.Value;

        }


    }
}
