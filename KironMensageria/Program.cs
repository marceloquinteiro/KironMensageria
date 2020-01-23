using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


namespace KironMensageria
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var email = new AutomaticEmail();
            int selector = 7;

            Console.WriteLine("****************************************************************");
            Console.WriteLine("Rotinas Kiron - Versão 1.01");
            Console.WriteLine();
            Console.WriteLine("****************************************************************");
            Console.WriteLine("Rotinas Disponíveis");
            Console.WriteLine("1 - Envio Diário de Carteira e Cota Líquida");
            Console.WriteLine("2 - Envio Semanal de Carteiras dos Fundos (Solicitação TAG/ITAU)");
            Console.WriteLine("3 - Rotina Alugueis");
            Console.WriteLine("4 - Stock Guide");
            Console.WriteLine("5 - BBG Data");
            Console.WriteLine("6 - Cancelar \n\n");
            while (selector != 1 && selector != 2 && selector != 3 && selector != 4 && selector != 5 && selector != 6)
            {
                Console.Write("Selecione a função desejada -> 1 / 2 / 3 / 4 / 5 / 6: ");
                Int32.TryParse(Console.ReadLine(), out selector);
            }
            Console.WriteLine("****************************************************************");

            switch (selector)
            {
                case 1:
                    email.FetchDiário();
                    Console.WriteLine("E-mails enviados com sucesso. Pressione qualquer tecla para encerrar");
                    Console.ReadKey();
                    break;

                case 2:
                    Console.Write("Indique o início do período no formato aaaammdd: ");
                    string inicio = Console.ReadLine();
                    Console.WriteLine();
                    Console.Write("Indique o fim do período no formato aaaammdd: ");
                    string fim = Console.ReadLine();
                    Console.WriteLine();

                    Console.WriteLine("1st day is {0} and last day is {1}", inicio, fim);
                    email.FetchQuinzenal(inicio, fim);
                    Console.WriteLine("E-mails enviados com sucesso. Pressione qualquer tecla para encerrar");
                    Console.ReadKey();
                    break;

                case 3:
                    email.CopiarCarteirasAlugueis();
                    Console.WriteLine("Carteiras atualizadas com sucesso. Pressione qualquer tecla para encerrar");
                    Console.ReadKey();
                    break;

                case 4:
                    var excel = new ExcelService();
                    excel.StockGuideUpdate();
                    break;

                case 5:
                    var excel1 = new ExcelService();
                    excel1.BbgStockGuidUpdate();
                    break;

                case 6:
                    break;
            }
        }
    }
}
