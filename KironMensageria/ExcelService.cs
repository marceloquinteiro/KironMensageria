using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;



namespace KironMensageria
{
    public static class TupleListExtensions
    {
        public static void Add<T1, T2>(this IList<Tuple<T1, T2>> list,
                T1 item1, T2 item2)
        {
            list.Add(Tuple.Create(item1, item2));
        }
    }

    public class ExcelService
    {
        private List<Tuple<string, string>> stocks = new List<Tuple<string, string>>() {
            {"Oi", "OIBR"},
            {"Ab-Inbev", "ABI"},
            {"ABC", "ABCB" },
            //{"Aliansce", "ALSC"},
            {"Alliar", "AALR"},
            {"Ambev", "ABEV"},
            {"Azul", "AZUL"},
            {"B3", "B3"},
            {"Banco do Brasil", "BBAS"},
            {"Bradesco", "BBDC"},
            {"BR Distribuidora", "BRDT" },
            {"BRFoods", "BRFS"},
            {"Brmalls", "BRML"},
            {"BTG Pactual", "BPAC" },
            {"Camil", "CAML"},
            {"Carrefour", "CRFB"},
            {"CCR", "CCRO"},
            {"Cemig", "CMIG"},
            {"Centauro", "CNTO" },
            {"Cosan", "CSAN"},
            {"CPFL", "CPFE"},
            {"CVC", "CVCB"},
            {"Cyrela", "CYRE"},
            {"Eneva", "ENEV"},
            {"Equatorial", "EQTL"},
            {"Even", "EVEN"},
            {"EzTec", "EZTC"},
            {"Fleury", "FLRY"},
            {"Gerdau", "GGBR"},
            {"Grupo Biotoscana", "GBIO"},
            {"Guararapes", "GUAR"},
            {"Hapvida", "HAPV"},
            {"Hering", "HGTX"},
            {"Hermes Pardini", "PARD"},
            {"Hypermarcas", "HYPE"},
            {"Iguatemi", "IGTA"},
            {"IMC", "MEAL"},
            {"Intermédica", "GNDI"},
            {"IRB", "IRBR"},
            {"Itau", "ITUB"},
            {"Klabin", "KLBN" },
            {"Linx", "LINX"},
            {"Localiza", "RENT"},
            {"Locamerica", "LCAM"},
            {"Lojas Americanas", "LAME"},
            {"Lopes", "LPSB"},
            {"M Dias Branco", "MDIA"},
            {"Magazine Luiza", "MGLU"},
            {"Mercado Livre", "MELI"},
            {"Mills", "MILS"},
            {"Movida", "MOVI"},
            {"Multiplan", "MULT"},
            {"Natura", "NATU"},
            {"Pao de acucar", "PCAR"},
            {"Petrobras", "PETR"},
            {"Qualicorp", "QUAL"},
            {"Raia Drogasil", "RADL"},
            {"Randon", "RAPT"},
            {"Renner", "LREN"},
            //{"Rodobens", "RDNI"},
            { "Romi", "ROMI" },
            {"Santander", "SANB"},
            {"Suzano", "SUZB"},
            {"Taesa", "TAEE"},
            {"Tegma", "TGMA"},
            {"TOTVS", "TOTS" },
            {"Ultrapar", "UGPA"},
            {"Vale", "VALE"},
            {"Via Varejo", "VVAR"},
            {"Weg", "WEGE"}
        };

        public void StockGuideUpdate()
        {
            Console.WriteLine("Operacional para {0} modelos", stocks.Count);
            var pathEnd = @"W:\Gestao\Market Data\Stock Guide\StockGuide.xlsm";
            var wkbEnd = OpenFile(pathEnd, false);
            var sheetController = 13;

            foreach (Tuple<string, string> modelo in stocks)
            {
                try
                {
                    string tupleName = modelo.Item1;
                    string tupleTicker = modelo.Item2;
                    var path = @"W:\Gestao\Empresas\" + tupleName + @"\Internal Analysis\" + tupleTicker + ".XLSX";
                    var wkb = OpenFile(path, true);
                    var stockGuideTable = FindTableWith(wkb, 1, "StockGuide_Table", "StockGuide_TableEnd");
                    CopyTableValues(stockGuideTable, wkbEnd, sheetController);
                    Clipboard.Clear();
                    wkb.Close(false);
                    sheetController++;
                }
                catch (Exception e)
                {
                    Console.WriteLine("Erro ao buscar as informações de {0}", modelo.Item2);
                    Console.WriteLine("Mensagem de erro: {0}", e);
                }
            }
            wkbEnd.Save();
            wkbEnd.Close();
            Console.Write("arquivos fechados. pressione qualquer tecla para encerrar");
            Console.ReadKey();
        }


        public void BbgStockGuidUpdate()
        {
            Console.WriteLine("Atualização de Banco de Dados Bloomberg");
            var pathEnd = @"W:\Gestao\Market Data\Stock Guide\StockGuide_BBGData.xlsx";
            var wkbEnd = OpenFile(pathEnd, false);
            var sheetController = 4;

            try
            {
                var path = @"W:\Gestao\Market Data\Stock Guide\Draft_Screen.xlsx";
                var wkb = OpenFile(path, true);
                var stockGuideTable = FindTableWith(wkb, 1, "Placeholder", "Placeholder");
                CopyTableValues(stockGuideTable, wkbEnd, sheetController);
                Clipboard.Clear();
                wkb.Close(false);
            }
            catch (Exception e)
            {
                Console.WriteLine("Erro ao buscar as informações do Banco de Dados Bloomberg");
                Console.WriteLine("Mensagem de erro: {0}", e);
            }

            wkbEnd.Save();
            wkbEnd.Close();
            Console.Write("arquivos fechados. pressione qualquer tecla para encerrar");
            Console.ReadKey();
        }

        private Excel.Workbook OpenFile(string path, bool _bool)
        {
            Excel.Application exc = new Excel.Application();
             return exc.Workbooks.Open(path, ReadOnly: _bool);
        }

        private Excel.Range FindTableWith(Excel.Workbook wkb, int sheetNo, string begin, string end)
        {
            // TODO: CHECK DE EXISTÊNCIA DE TABELA
            var sheets = wkb.Sheets;
            var activeSheet = sheets[sheetNo];
            Console.WriteLine(sheets[sheetNo].Name);
            var usedRange = (Excel.Range)activeSheet.UsedRange;
            //var controlCellBegin = (Excel.Range)usedRange.Find(begin);
            //var controlCellEnd = (Excel.Range)usedRange.Find(end);
            //return activeSheet.Range[controlCellBegin, controlCellEnd];
            return usedRange;
        }

        private void CopyTableValues(Excel.Range _rangeInput, Excel.Workbook _wkbEnd, int _sheetController)
        {
            _rangeInput.Copy();
            var sheetsEnd = _wkbEnd.Sheets;
            if (sheetsEnd.Count >= _sheetController)
            {
                Excel.Worksheet activeEnd = sheetsEnd[_sheetController];
                Excel.Range EndA1 = activeEnd.Range["A1"];
                EndA1.PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats);
            }
            else
            {
                Excel.Worksheet xlNewSheet = sheetsEnd.Add(sheetsEnd[_sheetController - 1]);
                var EndA1 = xlNewSheet.get_Range("a1");
                EndA1.PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats);
            }
        }
    }
}
