using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Vml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace bot_cep_correios
{
    class Program
    {

        static void Main(string[] args)
        {
            /*
             - Pegar dados da planilha do excel
             - Caminho do diretório até o arquivo do excel
            */
            var wbCep = new XLWorkbook(@"../../../../Lista_de_CEPs.xlsx");

            // Seleciona planilha a ser consumida
            var sheetWbCep = wbCep.Worksheet(1);

            
            // Caminho do diretório do arquivo resultado.xlsx que armazena os dados
            var wbResult = new XLWorkbook(@"../../../../resultado.xlsx");
            
            // Seleciona planilha para inserir os dados
            var sheetWbResult = wbResult.Worksheet(1);

            // Array de celulas para referencia de coluna
            string[] cells = new string[] { "A", "B", "C", "D", "E", "F"};

            // Linha inicial do primeiro registro da da planilha de CEP
            var lineWbCep = 2;

            // Linha inicial do primeiro registro que será inserido na planilha de Resultados
            var lineWbResult = 2;


            // Instanciando o ChromeDriver
            IWebDriver driver = new ChromeDriver();

            // Navegando até a URL para pegar os dados
            driver.Navigate().GoToUrl("http://www.buscacep.correios.com.br/sistemas/buscacep/BuscaCepEndereco.cfm");


            while (true)
            {
                // Pega o primeiro registo da tabela de CEP na coluna "B"
                var VeriryInitialCep = sheetWbCep.Cell("B" + lineWbCep.ToString()).Value.ToString();

                /*
                 - Verifica se existe valor lá dentro, caso não exista, ele para o laço e a aplicação é encerrada
                 - Essa é a ultima etapa do processo, que é quando acabam os registros de CEP
                */
                if (string.IsNullOrEmpty(VeriryInitialCep)) break;

                // Cep Inicial
                var InitialCep = Convert.ToInt32(VeriryInitialCep);

                // Cep Final
                var FinalCep = Convert.ToInt32(sheetWbCep.Cell("C" + lineWbCep.ToString()).Value);

                // While que vai do cepInicial até o cepFinal da linha atual
                while (InitialCep <= FinalCep)
                {

                    // Preenchendo valores no formulario
                    var cepInput = driver.FindElement(By.Name("relaxation"));
                    cepInput.SendKeys(InitialCep.ToString());

                    // Enviando valores
                    var submitButton = driver.FindElement(By.XPath("//input[@value='Buscar']"));
                    submitButton.Click();

                    // Elemento para fazer uma nova consulta
                    var newConsult = driver.FindElement(By.XPath("//span[@class='mver']/ul/li/a[@href='sistemas/buscacep/buscaCepEndereco.cfm']"));


                    // Tratando exceções
                    try
                    {
                        // Pega os registros obtidos no site dos correios e armazena nas colunas da planilha
                        for(int i = 0; i < 4; i++)
                        {
                            IWebElement td = driver.FindElement(By.XPath("//table[@class='tmptabela']/tbody/tr/td[" + (i + 1) + "]"));
                            string tdText = td.Text;

                            sheetWbResult.Cell(cells[i] + lineWbResult.ToString()).Value = tdText.ToString();
                        }

                        // Pega a data e a hora atual
                        DateTime processeDate = DateTime.Now;

                        // Armazena a data e a hora na planilha
                        sheetWbResult.Cell(cells[4] + lineWbResult.ToString()).Value = processeDate.ToString("dd/MM/yyyy");
                        sheetWbResult.Cell(cells[5] + lineWbResult.ToString()).Value = processeDate.ToString("HH:mm");

                        // Salva os dados na planilha
                        wbResult.Save();

                        // Incrementa +1 na faixa de CEP
                        InitialCep++;

                        // Incrementa +1, para poder inserir dados na linha de baixo do ultimo registro inserido
                        lineWbResult++;

                        // Fazer uma nova consulta
                        newConsult.Click();

                    }
                    // Se ele não achar a tag <table>, ele cai no "catch"
                    catch
                    {
                        // Incrementa +1 na faixa de CEP
                        InitialCep++;

                        // Fazer uma nova consulta
                        newConsult.Click();
                    }

                }
                // Incrementa +1 para ir na linha de baixo e pegar outra faixa de CEP
                lineWbCep++;

            }
        }
    }
}
