using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace KironMensageria
{
    class AutomaticEmail
    {
        /* Método compartilhado entre os demais, que envia os emails, conforme título, destinatários, corpo de texto e anexos desejados */
        public void SendEmail(string title, string[] recipients, string body, string[] attachments = null)
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem);
            mail.Subject = title;
            Outlook.AddressEntry currentUser = app.Session.CurrentUser.AddressEntry;

            if (currentUser.Type == "EX")
            {
                Outlook.ExchangeUser manager = currentUser.GetExchangeUser().GetExchangeUserManager();

                foreach (var nome in recipients)
                {
                    mail.Recipients.Add(nome);
                }

                mail.Recipients.ResolveAll();
                mail.HTMLBody = body + currentUser.Name.ToString() + "</ body ></ html >";

                if (attachments != null)
                {
                    foreach (var atch in attachments)
                    {
                        mail.Attachments.Add(atch,
                        Outlook.OlAttachmentType.olByValue, Type.Missing,
                        Type.Missing);
                    }

                }

                mail.Send();
            }
        }

        /* Esse método captura as carteiras dos fundos Kiron, para um período especificado pelo usuário. 
         * Sua utilização tem por objetivo encaminhar arquivos XML e XLS para o Itaú e TAG, respectivamente,
         * com 15 dias de delay */
        public void FetchQuinzenal(string _1st, string last)
        {
            List<string> xlsSelected = new List<string>();
            List<string> xmlSelected = new List<string>();
            List<string> DataInvoker = new List<string>();

            for (int date = Int32.Parse(_1st); date <= Int32.Parse(last); date++)
            {
                // Ajuste de Data para os e-mails
                string year = date.ToString().Substring(0, 4);
                string month = date.ToString().Substring(4, 2);
                string day = date.ToString().Substring(6, 2);
                string dateHuman = day + "." + month + "." + year;
                DataInvoker.Add(dateHuman);

                string[] xlsDirectories =
                    {
                    @"W:\Operacional\Fundos\Kiron I FIC FIA\Carteiras",
                    @"W:\Operacional\Fundos\Kiron II FIC FIA\Carteiras",
                    @"W:\Operacional\Fundos\Kiron III FIC FIA\Carteiras",
                    @"W:\Operacional\Fundos\Kiron Master FIA\Carteiras",
                    };

                string[] xmlDirectories =
                    {
                    @"W:\Operacional\Fundos\Kiron I FIC FIA\XML Anbima",
                    @"W:\Operacional\Fundos\Kiron II FIC FIA\XML Anbima",
                    @"W:\Operacional\Fundos\Kiron III FIC FIA\XML Anbima",
                    @"W:\Operacional\Fundos\Kiron Master FIA\XML Anbima",
                    @"W:\Operacional\Fundos\Kiron Institucional FIA\XML Anbima"
                    };

                foreach (var xlsdir in xlsDirectories)
                {
                    DirectoryInfo xlsdirectory = new DirectoryInfo(xlsdir);
                    string[] xlsfileNameArray =
                    {
                        "TAG CHIRON_" + date + ".xls",
                        "KIRON II_" + date + ".xls",
                        "KIRON III_" + date + ".xls",
                        "KIRON MST_" + date + ".xls",
                    };

                    foreach (var name in xlsfileNameArray)
                    {
                        try
                        {
                            FileInfo[] file = xlsdirectory.GetFiles(name);
                            xlsSelected.Add(file[0].FullName);
                        }
                        catch (Exception)
                        {
                            // do nothing
                        }
                    }
                }

                foreach (var xmldir in xmlDirectories)
                {
                    DirectoryInfo xmldirectory = new DirectoryInfo(xmldir);
                    string[] xmlfileNameArray =
                    {
                        "FD25213366000170_" + date + "*",
                        "FD28626275000154_" + date + "*",
                        "FD34475542000132_" + date + "*",
                        "FD28408139000198_" + date + "*",
                        "FD29054793000103_" + date + "*"
                    };

                    foreach (var name in xmlfileNameArray)
                    {
                        try
                        {
                            FileInfo[] file = xmldirectory.GetFiles(name);
                            xmlSelected.Add(file[0].FullName);
                        }
                        catch (Exception)
                        {
                            // do nothing;
                        }
                    }
                }
            }

            Console.WriteLine("Carteiras em XLS identificadas:");
            foreach (var item in xlsSelected)
            {
                Console.WriteLine(item);
            }
            Console.WriteLine();

            Console.WriteLine("Carteiras em XML identificadas:");
            foreach (var item in xmlSelected)
            {
                Console.WriteLine(item);
            }

            // Preparando o e-mail:

            // Anexos para envio no e-mail
            string[] attachmentTAG = xlsSelected.ToArray();
            string[] attachmentItau = xmlSelected.ToArray();

            // Ajuste de data para os e-mails
            string[] emailDayReference = DataInvoker.ToArray();

            // Títulos de e-mail:
            string titleItau = "Kiron - Solicitação de XML";
            string titleTAG = "Kiron - Carteiras atualizadas";

            // Destinatários de e-mail:
            //string[] emailListItau = { "caio.castro@kironcapital.com.br" }; // PARA TESTE 
            //string[] emailListTAG = { "caio.castro@kironcapital.com.br" }; // PARA TESTE
            string[] emailListXML = { "leticia-ramos.silva@itau-unibanco.com.br", "itau_risco@itau-unibanco.com.br", "carteiras@aditusbr.com", "atendimento.middle@itau-unibanco.com.br", "risco@taginvest.com.br", "administracao@kironcapital.com.br" };
            string[] emailListExcel = { "gestao@taginvest.com.br", "marco.bismarchi@taginvest.com.br", "bo@taginvest.com.br", "administracao@kironcapital.com.br" };

            // Corpos de texto padrão:
            string bodyItau = "<html><header><style> body { font-family:'Helvetica light'; font-size:11pt; } </style></header>" +
                "<body> Prezados, <br><br>" +
                "Seguem os XML para os nossos fundos.<br><br>" +
                "Um abraço,<br>";

            string bodyTAG = "<html><header><style> body { font-family:'Helvetica light'; font-size:11pt; } </style></header>" +
                 "<body> Caros, <br><br>" +
                 "Seguem anexas as carteiras dos nossos fundos.<br><br>" +
                 "Um abraço,<br>";

            SendEmail(titleItau, emailListXML, bodyItau, attachmentItau);
            Console.WriteLine("E-mail Itau enviado com sucesso.");
            SendEmail(titleTAG, emailListExcel, bodyTAG, attachmentTAG);
            Console.WriteLine("E-mail TAG enviado com sucesso.");

        }

        /* Esse método busca automaticamente o valor da cota líquida dos fundos Kiron Master FIA, Kiron FIC FIA, Kiron II FIC FIA e Kiron Allure FIA
         * referente ao último pregão e os envia a uma lista pré-determinada de destinatários */
        public void FetchDiário()

        {
            // especifica a data da carteira mais recente
            var lastBusinessDay = LastBusinessDay();

            // captura a carteira do Kiron FIC FIA
            Excel.Application kironFIC = new Excel.Application();
            string carteiraKironFIC = @"W:\Operacional\Fundos\Kiron I FIC FIA\Carteiras\TAG CHIRON_" + lastBusinessDay + ".XLS";
            string analiseKironFIC = @"W:\Operacional\Fundos\Kiron I FIC FIA\Carteiras\Analise_TAG CHIRON_" + lastBusinessDay + ".XLS";
            Excel.Workbook kironFICWb = kironFIC.Workbooks.Open(carteiraKironFIC);
            Excel.Sheets kironFICSheets = kironFICWb.Worksheets;
            Excel.Worksheet kironFICActiveSheet = kironFICSheets[1];
            Excel.Range rangeFIC = kironFICActiveSheet.UsedRange;
            int lastRowFIC = rangeFIC.Rows.Count;
            int selectedCollumFIC = 3;
            Excel.Range cellFIC = kironFICActiveSheet.Cells[lastRowFIC, selectedCollumFIC];
            double cotaLiquidaFIC = cellFIC.Value;
            kironFICWb.Close();
            Console.WriteLine("Kiron FIC FIA. Data: {0}. Cota Líquida: {1}", lastBusinessDay, cotaLiquidaFIC);

            // captura a carteira do Kiron II FIC FIA
            Excel.Application kironIIFIC = new Excel.Application();
            string carteiraKironIIFIC = @"W:\Operacional\Fundos\Kiron II FIC FIA\Carteiras\KIRON II_" + lastBusinessDay + ".XLS";
            string analiseKironIIFIC = @"W:\Operacional\Fundos\Kiron II FIC FIA\Carteiras\Analise_KIRON II_" + lastBusinessDay + ".XLS";
            Excel.Workbook kironIIFICWb = kironIIFIC.Workbooks.Open(carteiraKironIIFIC);
            Excel.Sheets kironIIFICSheets = kironIIFICWb.Worksheets;
            Excel.Worksheet kironIIFICActiveSheet = kironIIFICSheets[1];
            Excel.Range rangeIIFIC = kironIIFICActiveSheet.UsedRange;
            int lastRowIIFIC = rangeIIFIC.Rows.Count;
            int selectedCollumIIFIC = 3;
            Excel.Range cellIIFIC = kironIIFICActiveSheet.Cells[lastRowIIFIC, selectedCollumIIFIC];
            double cotaLiquidaIIFIC = cellIIFIC.Value;
            kironIIFICWb.Close();
            Console.WriteLine("Kiron II FIC FIA. Data: {0}. Cota Líquida: {1}", lastBusinessDay, cotaLiquidaIIFIC);

            // captura a carteira do Kiron III FIC FIA
            Excel.Application kironIIIFIC = new Excel.Application();
            string carteiraKironIIIFIC = @"W:\Operacional\Fundos\Kiron III FIC FIA\Carteiras\KIRON III_" + lastBusinessDay + ".XLS";
            string analiseKironIIIFIC = @"W:\Operacional\Fundos\Kiron III FIC FIA\Carteiras\Analise_KIRON III_" + lastBusinessDay + ".XLS";
            Excel.Workbook kironIIIFICWb = kironIIIFIC.Workbooks.Open(carteiraKironIIIFIC);
            Excel.Sheets kironIIIFICSheets = kironIIIFICWb.Worksheets;
            Excel.Worksheet kironIIIFICActiveSheet = kironIIIFICSheets[1];
            Excel.Range rangeIIIFIC = kironIIIFICActiveSheet.UsedRange;
            int lastRowIIIFIC = rangeIIIFIC.Rows.Count;
            int selectedCollumIIIFIC = 3;
            Excel.Range cellIIIFIC = kironIIIFICActiveSheet.Cells[lastRowIIIFIC, selectedCollumIIIFIC];
            double cotaLiquidaIIIFIC = cellIIIFIC.Value;
            kironIIIFICWb.Close();
            Console.WriteLine("Kiron III FIC FIA. Data: {0}. Cota Líquida: {1}", lastBusinessDay, cotaLiquidaIIIFIC);

            // captura a carteira do Kiron MST
            Excel.Application kironMST = new Excel.Application();
            string carteiraKironMST = @"W:\Operacional\Fundos\Kiron Master FIA\Carteiras\KIRON MST_" + lastBusinessDay + ".XLS";
            string analiseKironMST = @"W:\Operacional\Fundos\Kiron Master FIA\Carteiras\Analise_KIRON MST_" + lastBusinessDay + ".XLS";
            Excel.Workbook kironMSTWb = kironMST.Workbooks.Open(carteiraKironMST);
            Excel.Sheets kironMSTSheets = kironMSTWb.Worksheets;
            Excel.Worksheet kironMSTActiveSheet = kironMSTSheets[1];
            Excel.Range rangeMST = kironMSTActiveSheet.UsedRange;
            int lastRowMST = rangeMST.Rows.Count;
            int selectedCollumMST = 3;
            Excel.Range cellMST = kironMSTActiveSheet.Cells[lastRowMST, selectedCollumMST];
            double cotaLiquidaMST = cellMST.Value;
            kironMSTWb.Close();
            Console.WriteLine("Kiron Master FIA. Data: {0}. Cota Líquida: {1}", lastBusinessDay, cotaLiquidaMST);

            // captura a carteira do Kiron Allure FIA
            Excel.Application kironAllure = new Excel.Application();
            string carteiraKironAllure = @"W:\Operacional\Fundos\Kiron Allure\Carteiras\KIRON ALLURE_" + lastBusinessDay + ".XLS";
            string analiseKironAllure = @"W:\Operacional\Fundos\Kiron Allure\Carteiras\Analise_KIRON ALLURE_" + lastBusinessDay + ".XLS";
            Excel.Workbook kironAllureWb = kironAllure.Workbooks.Open(carteiraKironAllure);
            Excel.Sheets kironAllureSheets = kironAllureWb.Worksheets;
            Excel.Worksheet kironAllureActiveSheet = kironAllureSheets[1];
            Excel.Range rangeAllure = kironAllureActiveSheet.UsedRange;
            int lastRowAllure = rangeAllure.Rows.Count;
            int selectedCollumAllure = 3;
            Excel.Range cellAllure = kironAllureActiveSheet.Cells[lastRowAllure, selectedCollumAllure];
            double cotaLiquidaAllure = cellAllure.Value;
            kironAllureWb.Close();
            Console.WriteLine("Kiron Allure FIA. Data: {0}. Cota Líquida: {1}", lastBusinessDay, cotaLiquidaAllure);

            // captura a carteira do Kiron Institucional FIA
            Excel.Application kironInst = new Excel.Application();
            string carteiraKironInst = @"W:\Operacional\Fundos\Kiron Institucional FIA\Carteiras\KIRON INST_" + lastBusinessDay + ".XLS";
            string analiseKironInst = @"W:\Operacional\Fundos\Kiron Institucional FIA\Carteiras\Analise_KIRON INST_" + lastBusinessDay + ".XLS";
            Excel.Workbook kironInstWb = kironInst.Workbooks.Open(carteiraKironInst);
            Excel.Sheets kironInstSheets = kironInstWb.Worksheets;
            Excel.Worksheet kironInstActiveSheet = kironInstSheets[1];
            Excel.Range rangeInst = kironInstActiveSheet.UsedRange;
            int LastRowInst = rangeInst.Rows.Count;
            int selectedCollumInst = 3;
            Excel.Range cellInst = kironInstActiveSheet.Cells[LastRowInst, selectedCollumInst];
            double cotaLiquidaInst = cellInst.Value;
            kironInstWb.Close();
            Console.WriteLine("Kiron Institucional FIA. Data: {0}. Cota Líquida: {1}", lastBusinessDay, cotaLiquidaInst);

            Console.WriteLine("Confirma Cotas? (Y/N)");
            var check = Console.ReadLine();
            if (check != "Y" && check != "y")
            {
                Environment.Exit(0);
            }

            // Anexos para envio no e-mail
            string[] attachmentCarteiras = {carteiraKironFIC,carteiraKironIIFIC, carteiraKironIIIFIC, carteiraKironMST, carteiraKironAllure, analiseKironFIC,
                analiseKironIIFIC, analiseKironIIIFIC, analiseKironMST, analiseKironAllure, carteiraKironInst, analiseKironInst};

            // Ajuste de Data para os e-mails
            string year = lastBusinessDay.Substring(0, 4);
            string month = lastBusinessDay.Substring(4, 2);
            string day = lastBusinessDay.Substring(6, 2);
            string emailDayReference = day + "." + month + "." + year;

            // Títulos de e-mail:
            string titleCotas = "Kiron - Cota diária fundos";
            string titleCarteirasDiarias = "Carteiras Kiron - " + emailDayReference;

            // Destinatários de e-mail:
            //string[] emailListCotas = { "caio.castro@kironcapital.com.br" }; // PARA TESTE
            //string[] emailListCarteirasDiarias = { "caio.castro@kironcapital.com.br" }; // PARA TESTE
            string[] emailListCotas = { "OL-cotas@btgpactual.com", "cotas.externas-itau@itau-unibanco.com.br",
                "4010.cadproprios@bradesco.com.br", "cotas@orama.com.br", "bo@taginvest.com.br", "administracao@kironcapital.com.br",
                "fof@itau-unibanco.com.br", "dilton@borduna.com.br", "backoffice@sjasset.com.br", "ef@armoryinvest.com",
                "OL-Cotas-Fundos-Digital@btgpactual.com", "eduardo.araujo@deminvest.com.br", "tamiris@deminvest.com.br", "luiza@deminvest.com.br",
                "nucleoproc@bradesco.com.br", "nucleo12@bradesco.com.br", "dac.cotas@bradesco.com.br", "bo-is@vincipartners.com","cota@genialinvestimentos.com.br",
                "infos.asset@safra.com.br"
            };
            string[] emailListCarteirasDiarias = { "administracao@kironcapital.com.br", "francisco.utsch@kironcapital.com.br",
                "luiz.liuzzi@kironcapital.com.br" };

            // Corpos de texto padrão:
            string bodyCotas = "<html><header><style> body { font-family:'Helvetica light'; font-size:11pt; } </style></header>" +
                "<body> Caros,<br><br>" +
                "Carteiras do dia " + emailDayReference +
                "<br><br>Kiron FIC FIA:<br>" +
                " - " + cotaLiquidaFIC.ToString() +
                "<br><br>Kiron II FIC FIA:<br>" +
                " - " + cotaLiquidaIIFIC.ToString() +
                "<br><br>Kiron III FIC FIA:<br>" +
                " - " + cotaLiquidaIIIFIC.ToString() +
                "<br><br>Kiron Master FIA:<br>" +
                " - " + cotaLiquidaMST.ToString() +
                "<br><br>Kiron Institucional FIA:<br>" +
                " - " + cotaLiquidaInst.ToString() +
                "<br><br>Um abraço, <br>";

            string bodyCarteirasDiárias = "<html><header><style> body { font-family:'Helvetica light'; font-size:11pt; } </style></header>" +
                "<body> Caros,<br><br>" +
                "Carteiras do dia " + emailDayReference +
                "<br><br>Abs. <br>";

            SendEmail(titleCarteirasDiarias, emailListCarteirasDiarias, bodyCarteirasDiárias, attachmentCarteiras);
            Console.WriteLine("E-mail Carteiras Diárias enviado com sucesso.");
            SendEmail(titleCotas, emailListCotas, bodyCotas);
            Console.WriteLine("E-mail Cotas enviado com sucesso.");

            // Salva a carteira do KironMST no destino @"W:\Operacional\Cart Hoje"
            AtualizaCarteiraDiaria();

        }

        /* Identifica a última carteira disponível no FileSite e pede por confirmação do usuário */
        public string LastBusinessDay()
        {
            bool i = true;
            int counter = -1;
            while (i == true)
            {
                string result = DateTime.Today.AddDays(counter).ToString("yyyyMMdd");
                string carteiraKironMST = @"W:\Operacional\Fundos\Kiron Master FIA\Carteiras\KIRON MST_" + result + ".XLS";
                if (File.Exists(carteiraKironMST))
                {
                    i = false;
                    Console.WriteLine("Última carteira válida: {0}. \nConfirma? (Y/N)", carteiraKironMST);
                    string response = Console.ReadLine();
                    if (response == "Y" || response == "y")
                    {
                        return result;
                    }
                    Environment.Exit(0);
                }
                counter--;
            }
            return "";
        }

        /* Identifica a última carteira disponível no FileSite sem pedir confirmação do usuário */
        public string LastBusinessDay(bool confirmationcheck)
        {
            bool i = true;
            int counter = -1;
            while (i == true)
            {
                string result = DateTime.Today.AddDays(counter).ToString("yyyyMMdd");
                string carteiraKironMST = @"W:\Operacional\Fundos\Kiron Master FIA\Carteiras\KIRON MST_" + result + ".XLS";
                if (File.Exists(carteiraKironMST))
                {
                    i = false;
                    if (confirmationcheck == false)
                    {
                        return result;
                    }

                    else
                    {
                        Console.WriteLine("Última carteira válida: {0}. \nConfirma? (Y/N)", carteiraKironMST);
                        string response = Console.ReadLine();
                        if (response == "Y" || response == "y")
                        {
                            return result;
                        }
                        Environment.Exit(0);
                    }
                }
                counter--;
            }
            return "";
        }

        /* */
        public void AtualizarCarteiras()
        {
            // substitui cópia do modelo - Vale

            DirectoryInfo ValeDirectory = new DirectoryInfo(@"W:\Gestao\Market Data\Suporte");

            try
            {
                // deleta e copia
                FileInfo[] oldFile = ValeDirectory.GetFiles("VALE.1.xlsx");
                oldFile[0].Delete();
                DirectoryInfo ValeOriginalDirectory = new DirectoryInfo(@"W:\Gestao\Empresas\Vale\Internal Analysis");
                FileInfo[] file = ValeOriginalDirectory.GetFiles("VALE.xlsx");
                file[0].CopyTo(@"W:\Gestao\Market Data\Suporte\VALE.1.xlsx");
            }
            catch (Exception)
            {
                // só copia
                DirectoryInfo ValeOriginalDirectory = new DirectoryInfo(@"W:\Gestao\Empresas\Vale\Internal Analysis");
                FileInfo[] file = ValeOriginalDirectory.GetFiles("VALE.xlsx");
                file[0].CopyTo(@"W:\Gestao\Market Data\Suporte\VALE.1.xlsx");
            }

            // substitui cópia do modelo - Gerdau

            DirectoryInfo GerdauDirectory = new DirectoryInfo(@"W:\Gestao\Market Data\Suporte");

            try
            {
                // deleta e copia
                FileInfo[] oldFile = GerdauDirectory.GetFiles("GGBR.1.xlsx");
                oldFile[0].Delete();
                DirectoryInfo GerdauOriginalDirectory = new DirectoryInfo(@"W:\Gestao\Empresas\Gerdau\Internal Analysis");
                FileInfo[] file = GerdauOriginalDirectory.GetFiles("GGBR.xlsx");
                file[0].CopyTo(@"W:\Gestao\Market Data\Suporte\GGBR.1.xlsx");
            }
            catch (Exception)
            {
                // só copia
                DirectoryInfo GerdauOriginalDirectory = new DirectoryInfo(@"W:\Gestao\Empresas\Gerdau\Internal Analysis");
                FileInfo[] file = GerdauOriginalDirectory.GetFiles("GGBR.xlsx");
                file[0].CopyTo(@"W:\Gestao\Market Data\Suporte\GGBR.1.xlsx");


            }

            // substitui cópia do modelo - Itaú

            DirectoryInfo ItauDirectory = new DirectoryInfo(@"W:\Gestao\Market Data\Suporte");

            try
            {
                // deleta e copia
                FileInfo[] oldFile = ItauDirectory.GetFiles("ITUB.1.xlsx");
                oldFile[0].Delete();
                DirectoryInfo ItauOriginalDirectory = new DirectoryInfo(@"W:\Gestao\Empresas\Itau");
                FileInfo[] file = ItauOriginalDirectory.GetFiles("ITUB.xlsx");
                file[0].CopyTo(@"W:\Gestao\Market Data\Suporte\ITUB.1.xlsx");
            }
            catch (Exception)
            {
                // só copia
                DirectoryInfo ItauOriginalDirectory = new DirectoryInfo(@"W:\Gestao\Empresas\Itau");
                FileInfo[] file = ItauOriginalDirectory.GetFiles("ITUB.xlsx");
                file[0].CopyTo(@"W:\Gestao\Market Data\Suporte\ITUB.1.xlsx");

            }
        }

        /* */
        public void CopiarCarteirasAlugueis()
        {
            // especifica a data da carteira mais recente
            var lastBusinessDay = LastBusinessDay();

            // captura a carteira do Kiron MST
            Excel.Application kironMST = new Excel.Application();
            string carteiraKironMST = @"W:\Operacional\Fundos\Kiron Master FIA\Carteiras\KIRON MST_" + lastBusinessDay + ".XLS";
            Excel.Workbook kironMSTWb = kironMST.Workbooks.Open(carteiraKironMST);
            Excel.Sheets kironMSTSheets = kironMSTWb.Worksheets;
            Excel.Worksheet kironMSTActiveSheet = kironMSTSheets[1];
            Excel.Range novoCellsMST = kironMSTActiveSheet.Cells;


            // captura a carteira do Kiron Institucional FIA
            Excel.Application kironInst = new Excel.Application();
            string carteiraKironInst = @"W:\Operacional\Fundos\Kiron Institucional FIA\Carteiras\KIRON INST_" + lastBusinessDay + ".XLS";
            Excel.Workbook kironInstWb = kironInst.Workbooks.Open(carteiraKironInst);
            Excel.Sheets kironInstSheets = kironInstWb.Worksheets;
            Excel.Worksheet kironInstActiveSheet = kironInstSheets[1];
            Excel.Range novoCellsInst = kironInstActiveSheet.Cells;


            // captura a carteira destino e seleciona a aba correta de cola

            // Excel.Range velhoCellsMST
            Excel.Application kironAlug = new Excel.Application();
            string alugueisKiron = @"W:\Gestao\Market Data\Aluguel de Acoes\2018.03.15 Aluguel de Acoes.xlsm";
            Excel.Workbook kironAlugWb = kironAlug.Workbooks.Open(alugueisKiron);
            Excel.Sheets kironAlugSheets = kironAlugWb.Worksheets;

            // Loop MST Inst e GP
            Excel.Worksheet kironAlugMSTActiveSheet = kironAlugSheets[3];
            Excel.Range velhoCellsMST = kironAlugMSTActiveSheet.Cells;
            kironMSTActiveSheet.UsedRange.Copy();
            var kironMSTa1 = kironAlugMSTActiveSheet.get_Range("a1");
            kironMSTa1.PasteSpecial(Excel.XlPasteType.xlPasteValues);

            Excel.Worksheet kironAlugInstActiveSheet = kironAlugSheets[4];
            Excel.Range velhoCellsInst = kironAlugInstActiveSheet.Cells;
            kironInstActiveSheet.UsedRange.Copy();
            var kironInsta1 = kironAlugInstActiveSheet.get_Range("a1");
            kironInsta1.PasteSpecial(Excel.XlPasteType.xlPasteValues);


            // fecha Inst

            Clipboard.Clear();
            kironInstWb.Close();
            kironMSTWb.Close();
            kironAlugWb.Save();
            kironAlugWb.Close();

        }

        /* rotina que copia a carteira do kiron MST para a pasta solicitada pelo Liuzzi */
        public void AtualizaCarteiraDiaria()
        {
            var lastBusinessDay = LastBusinessDay(false);

            DirectoryInfo liuzziDirectory = new DirectoryInfo(@"W:\Operacional\Cart Hoje");
            FileInfo[] oldFile = liuzziDirectory.GetFiles("h.XLS");
            oldFile[0].Delete();
            DirectoryInfo xlsMSTDirectory = new DirectoryInfo(@"W:\Operacional\Fundos\Kiron Master FIA\Carteiras");
            FileInfo[] file = xlsMSTDirectory.GetFiles("KIRON MST_" + lastBusinessDay + ".XLS");
            file[0].CopyTo(@"W:\Operacional\Cart Hoje\h.xls");
            Console.WriteLine("Cart Hoje atualizado");
        }
    }
}
