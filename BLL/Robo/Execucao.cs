using BLL.Entity;
using BLL.Services;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using VersaoChromedriver;

namespace BLL.Robo
{
    public class Execucao
    {
        public void SolicitarPagamento(string Usuario, string Senha, string Data, List<BaixarEntity> dadosPagamento)
        {

            ChromeOptions options = new ChromeOptions();

            options.AddArgument("--start-maximized");
            //options.AddUserProfilePreference("download.defaut_directory", caminho);

            versionador v = new versionador();
            var chromeDriverService = ChromeDriverService.CreateDefaultService(v.versaoChromeDriver());
            chromeDriverService.HideCommandPromptWindow = true;

            using (ChromeDriver driver = new ChromeDriver(chromeDriverService, options: options))
            {
                try
                {
                    driver.Navigate().GoToUrl("https://mercadolivre.elaw.com.br/#");

                }
                catch (Exception)
                {
                    return;
                }

                driver.FindElement(By.Name("username")).SendKeys(Usuario);
                driver.FindElement(By.Name("password")).SendKeys(Senha);

                driver.FindElement(By.LinkText("Acessar")).Click();
                Thread.Sleep(10000);

                foreach (BaixarEntity pagamento in dadosPagamento)
                {
                    driver.Navigate().GoToUrl("https://mercadolivre.elaw.com.br/processoView.elaw");
                    try
                    {
                        //driver.SwitchTo().DefaultContent();
                        Waitforload(driver);
                        Thread.Sleep(1000);

                        while (confereitem(driver, "//input[@placeholder='Pesquise por aqui!']"))
                        {

                        }

                        driver.FindElement(By.XPath("//input[@placeholder='Pesquise por aqui!']")).Clear();
                        Thread.Sleep(1000);
                        Waitforload(driver);
                        driver.FindElement(By.XPath("//input[@placeholder='Pesquise por aqui!']")).SendKeys(pagamento.Pasta_Cliente);
                        

                        Thread.Sleep(3000);

                        driver.SwitchTo().DefaultContent();

                        //verifica o conteúdo do autocompletar e clica --
                        var tempo = DateTime.Now;
                        try
                        {
                            while (ConfereitemClass(driver, "ui-autocomplete-query"))
                            {
                                if (DateTime.Now.AddSeconds(-30) >= tempo)
                                {
                                    throw new Exception("Tempo limite atingido, verificar pasta");
                                }

                            }

                        }
                        catch (Exception e)
                        {
                            GravarLog(pagamento,"Tempo limite atingido, verificar pasta");

                            EscreveExcel novo1 = new EscreveExcel();
                            pagamento.Status = "Tempo limite atingido, verificar pasta";
                            novo1.GeraLogXLSX(dadosPagamento);
                            continue;
                        }

                        var op = driver.FindElements(By.ClassName("ui-autocomplete-query"));


                        foreach (var iop in op)
                        {
                            if (iop.Text.Contains(pagamento.Pasta_Cliente))
                            {
                                iop.Click();

                                Waitforload(driver);
                                break;
                            }
                        }

                        while (ConfereitemLinktext(driver, "Pagamentos"))
                        {

                        }

                        pagamento.Proc_Jud = driver.FindElement(By.Id("j_id_hd:j_id_hf_7_4_2s_1_2_1:j_id_hf_7_4_2s_1_2_2_1_2_1:j_id_hf_7_4_2s_1_2_2_1_2_2_1_1")).Text;

                        string statusDoProcesso = driver.FindElement(By.XPath("//*[@id='processoDadosCabecalhoForm']/table/tbody/tr[5]/td[6]/label")).Text;

                        if (!statusDoProcesso.Equals("Ativo"))
                        {
                            EscreveExcel novoItem2 = new EscreveExcel();
                            pagamento.Status = "ENCERRADO";
                            pagamento.diaHora = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
                            novoItem2.GeraLogXLSX(dadosPagamento);
                            GravarLog(pagamento, "ENCERRADO");
                            continue;
                        }

                        if (pagamento.Tipo_Pagamento.ToUpper().Contains("ACORDO") || pagamento.Tipo_Pagamento.ToUpper().Contains("CONDENAÇÃO"))
                        {
                            EscreveExcel novoItem2 = new EscreveExcel();
                            pagamento.Status = "Robô não habilitado para solicitar pagamentos deste tipo! Sem ações a tomar.";
                            novoItem2.GeraLogXLSX(dadosPagamento);
                            GravarLog(pagamento, "Robô não habilitado para solicitar pagamentos deste tipo! Sem ações a tomar.");
                            continue;
                        }

                        Thread.Sleep(2000);
                        Waitforload(driver);

                        driver.FindElement(By.LinkText("Pagamentos")).Click();
                        Waitforload(driver);

                        //botao novo pagamento
                        Thread.Sleep(1000);
                        driver.FindElement(By.Id("tabViewProcesso:j_id_i1_i_1_g_b")).Click();

                        Waitforload(driver);

                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_20_label")).Click();
                        Thread.Sleep(2000);

                        string contaContabil = "";
                        string centroCusto = "";
                        string textoModificado = "";

                        /*if (pagamento.Tipo_Pagamento.ToUpper().Contains("ACORDO"))
                        {
                            pagamento.Tipo_Pagamento = "Acordo";
                            contaContabil = "631011";
                            centroCusto = "BR - ACORDO";
                            textoModificado = "acordo adiantado";
                        }
                        if (pagamento.Tipo_Pagamento.ToUpper().Contains("CONDENAÇÃO"))
                        {
                            pagamento.Tipo_Pagamento = "Condenação";
                            contaContabil = "631013";
                            centroCusto = "BR - CONDENAÇÃO";
                            textoModificado = "condenação adiantado";
                        }*/

                        pagamento.Tipo_Pagamento = "Custas";
                        contaContabil = "631025";
                        centroCusto = "BR - CUSTAS PROCESSUAIS";
                        textoModificado = "custas adiantadas";


                        //Tipo de pagamento ok
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_20_filter")).SendKeys(pagamento.Tipo_Pagamento + Keys.Enter);
                        Thread.Sleep(2000);

                        //Valor ok
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:valorField_input")).Clear();
                        Thread.Sleep(1000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:valorField_input")).SendKeys(pagamento.Valor);
                        Thread.Sleep(2000);


                        //logica prazo fatal ok
                        //string data = verificaDia();
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_3_1_9_1q_1_1_1:pvpEFBfieldDate_input")).Click();
                        Thread.Sleep(500);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_3_1_9_1q_1_1_1:pvpEFBfieldDate_input")).SendKeys(Data);
                        Thread.Sleep(3000);
                        //driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:valorField_input")).Click();
                        //Thread.Sleep(1000);

                        //Data de Vencimento ok
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_4_1_9_7_1_input")).Click();
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_4_1_9_7_1_input")).SendKeys(Data);
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:valorField_input")).Click();
                        Thread.Sleep(1000);

                        //var listaLI = driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:processoValorFavorecido_panel")).FindElements(By.TagName("li")).ToList();

                        //Favorecido
                        var novoTempo = DateTime.Now;
                        try
                        {
                            //driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:processoValorFavorecido_input")).Clear();
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:processoValorFavorecido_input")).Clear();
                            Thread.Sleep(1000);
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:processoValorFavorecido_input")).SendKeys("GONDIM ADVOGADOS ASSOCIADOS");
                            Thread.Sleep(5000);
                            while (ConfereitemId(driver, "j_id_2w")) { }

                        }
                        catch (Exception e)
                        {
                            //ADICIONAR LOG
                            continue;
                        }

                        //ConfereitemClass(driver, "ui-autocomplete-query");
                         while (ConfereitemClass(driver, "ui-autocomplete-query"))
                         {
                             if (DateTime.Now.AddSeconds(-30) >= novoTempo)
                             {
                                throw new Exception("Não é possivel encontrar elemento");
                             }

                         }
                        var novoOp = driver.FindElements(By.ClassName("ui-autocomplete-query"));


                        foreach (var iop in novoOp)
                        {

                            Thread.Sleep(2000);
                            if (iop.Text.Contains("GONDIM ADVOGADOS ASSOCIADOS"))
                            {
                                iop.Click();

                                Waitforload(driver);
                                break;
                            }
                        }

                        //Tipo de Transferencia
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_6_1_9_1q_1_1_1:j_id_2l_28_6_1_9_1q_1_1_c_label")).Click();
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_6_1_9_1q_1_1_1:j_id_2l_28_6_1_9_1q_1_1_c_filter")).SendKeys("TED" + Keys.Enter);
                        Thread.Sleep(2000);

                        //Provedor
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_7_1_9_1q_1_1_1:j_id_2l_28_7_1_9_1q_1_1_c_label")).Click();
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_7_1_9_1q_1_1_1:j_id_2l_28_7_1_9_1q_1_1_c_filter")).SendKeys("9800000071" + Keys.Enter);
                        Thread.Sleep(2000);

                        //Tipo de Conta
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_8_1_9_1q_1_1_1:j_id_2l_28_8_1_9_1q_1_1_c_label")).Click();
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_8_1_9_1q_1_1_1:j_id_2l_28_8_1_9_1q_1_1_c_filter")).SendKeys("Checking" + Keys.Enter);
                        Thread.Sleep(2000);

                        //FECHA ESTADO
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_9_1_9_1q_1_1_1:pvpEFBfieldDate_input")).Click();
                        Thread.Sleep(500);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_9_1_9_1q_1_1_1:pvpEFBfieldDate_input")).SendKeys(Data);
                        Thread.Sleep(2000);

                        //Acordo
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_a_1_9_1q_1_1_1:j_id_2l_28_a_1_9_1q_1_1_c_label")).Click();
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_a_1_9_1q_1_1_1:j_id_2l_28_a_1_9_1q_1_1_c_filter")).SendKeys("JUDICIAL");
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_a_1_9_1q_1_1_1:j_id_2l_28_a_1_9_1q_1_1_c_2")).Click();

                        //Tipo de Procedimento
                        string id = "";
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_b_1_9_1q_1_1_1:j_id_2l_28_b_1_9_1q_1_1_c_label")).Click();
                        Thread.Sleep(2000);
                        
                        if (pagamento.Proc_Jud == "JEC")
                        {
                            id = "processoValorPagamentoEditForm:pvp:j_id_2l_28_b_1_9_1q_1_1_1:j_id_2l_28_b_1_9_1q_1_1_c_3";
                        }
                        else
                        {
                            id = "processoValorPagamentoEditForm:pvp:j_id_2l_28_b_1_9_1q_1_1_1:j_id_2l_28_b_1_9_1q_1_1_c_1";
                        }
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id(id)).Click();

                        //ContaContabil
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_c_1_9_1q_1_1_1:j_id_2l_28_c_1_9_1q_1_1_c_label")).Click();
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_c_1_9_1q_1_1_1:j_id_2l_28_c_1_9_1q_1_1_c_filter")).SendKeys(contaContabil);
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_c_1_9_1q_1_1_1:j_id_2l_28_c_1_9_1q_1_1_c_4")).Click();

                        //Centro de custo
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_d_1_9_1q_1_1_1:j_id_2l_28_d_1_9_1q_1_1_c_label")).Click();
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_d_1_9_1q_1_1_1:j_id_2l_28_d_1_9_1q_1_1_c_filter")).SendKeys(centroCusto);
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_d_1_9_1q_1_1_1:j_id_2l_28_d_1_9_1q_1_1_c_3")).Click();

                        //Forma de pagamento - deposito bancario
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:pvpEFSpgTypeSelectField1CombosCombo_label")).Click();
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:pvpEFSpgTypeSelectField1CombosCombo_filter")).SendKeys("Depósito Bancário");
                        Thread.Sleep(3000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:pvpEFSpgTypeSelectField1CombosCombo_3")).Click();
                        Thread.Sleep(1000);

                        //Dados bancário - Banco
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:j_id_2l_28_e_1_9_1q_1_2_c_2:j_id_2l_28_e_1_9_1q_1_2_c_7:0:j_id_2l_28_e_1_9_1q_1_2_c_10:j_id_2l_28_e_1_9_1q_1_2_c_az_label")).Click();
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:j_id_2l_28_e_1_9_1q_1_2_c_2:j_id_2l_28_e_1_9_1q_1_2_c_7:0:j_id_2l_28_e_1_9_1q_1_2_c_10:j_id_2l_28_e_1_9_1q_1_2_c_az_filter")).SendKeys("237"); //BRADESCO
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:j_id_2l_28_e_1_9_1q_1_2_c_2:j_id_2l_28_e_1_9_1q_1_2_c_7:0:j_id_2l_28_e_1_9_1q_1_2_c_10:j_id_2l_28_e_1_9_1q_1_2_c_az_155")).Click();//BRADESCO
                        Thread.Sleep(1000);

                        //Dados bancário - Agencia
                        try
                        {
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:j_id_2l_28_e_1_9_1q_1_2_c_2:j_id_2l_28_e_1_9_1q_1_2_c_7:0:j_id_2l_28_e_1_9_1q_1_2_c_10:bancarioAgencia")).Click();
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:j_id_2l_28_e_1_9_1q_1_2_c_2:j_id_2l_28_e_1_9_1q_1_2_c_7:0:j_id_2l_28_e_1_9_1q_1_2_c_10:bancarioAgencia")).SendKeys("0468");
                            Thread.Sleep(1000);
                        }
                        catch
                        {
                            Thread.Sleep(2000);
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:j_id_2l_28_e_1_9_1q_1_2_c_2:j_id_2l_28_e_1_9_1q_1_2_c_7:0:j_id_2l_28_e_1_9_1q_1_2_c_10:bancarioAgencia")).Click();
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:j_id_2l_28_e_1_9_1q_1_2_c_2:j_id_2l_28_e_1_9_1q_1_2_c_7:0:j_id_2l_28_e_1_9_1q_1_2_c_10:bancarioAgencia")).SendKeys("0468");
                            Thread.Sleep(1000);
                        }

                        //Dados bancário - Conta
                        try
                        {
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:j_id_2l_28_e_1_9_1q_1_2_c_2:j_id_2l_28_e_1_9_1q_1_2_c_7:0:j_id_2l_28_e_1_9_1q_1_2_c_10:bancarioContaCorrente")).Click();
                            Thread.Sleep(1000);
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:j_id_2l_28_e_1_9_1q_1_2_c_2:j_id_2l_28_e_1_9_1q_1_2_c_7:0:j_id_2l_28_e_1_9_1q_1_2_c_10:bancarioContaCorrente")).SendKeys("0627178-2");
                        }
                        catch
                        {
                            Thread.Sleep(2000);
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:j_id_2l_28_e_1_9_1q_1_2_c_2:j_id_2l_28_e_1_9_1q_1_2_c_7:0:j_id_2l_28_e_1_9_1q_1_2_c_10:bancarioContaCorrente")).Click();
                            Thread.Sleep(1000);
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_e_1_9_1q_1_2_1:j_id_2l_28_e_1_9_1q_1_2_c_2:j_id_2l_28_e_1_9_1q_1_2_c_7:0:j_id_2l_28_e_1_9_1q_1_2_c_10:bancarioContaCorrente")).SendKeys("0627178-2");
                        }

                        //Caminho do arquivo
                        try
                        {
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_f_1_9_e_3_input")).SendKeys(pagamento.Caminho_Arquivo);
                            Thread.Sleep(4000);
                        }
                        catch
                        {
                            string erroMsg = "ERRO: Não foi possível anexar o documento!";
                            GravarLog(pagamento,erroMsg);
                            EscreveExcel novoItem = new EscreveExcel();
                            pagamento.Status = erroMsg;
                            novoItem.GeraLogXLSX(dadosPagamento);
                            continue;
                        }

                        driver.ExecuteScript($"document.getElementById('processoValorPagamentoEditForm:pvp:gedEFileDataTable:0:j_id_2l_28_f_1_9_e_c_label').click()");
                        //driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:gedEFileDataTable:0:j_id_2l_28_f_1_9_e_c_label")).Click();
                        Thread.Sleep(3000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:gedEFileDataTable:0:j_id_2l_28_f_1_9_e_c_filter")).SendKeys("Comprovante de Pagamento");
                        Thread.Sleep(2000);
                        driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:gedEFileDataTable:0:j_id_2l_28_f_1_9_e_c_11")).Click();
                        Thread.Sleep(2000);

                        //Descrição
                        try
                        {
                            string parecer = $"Pagamento de {textoModificado}, assim solicitamos reembolso de pagamento ao escritório.";
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_g_1_9_3_1")).Click();
                            Thread.Sleep(1000);
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_g_1_9_3_1")).SendKeys(parecer);
                            Thread.Sleep(1000);
                        }
                        catch
                        {
                            Thread.Sleep(2000);
                            string parecer = $"Pagamento de {textoModificado}, assim solicitamos reembolso de pagamento ao escritório.";
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_g_1_9_3_1")).Click();
                            Thread.Sleep(1000);
                            driver.FindElement(By.Id("processoValorPagamentoEditForm:pvp:j_id_2l_28_g_1_9_3_1")).SendKeys(parecer);
                            Thread.Sleep(1000);
                        }

                        //botão SALVAR
                        try
                        {
                            driver.ExecuteScript($"document.getElementById('processoValorPagamentoEditForm:btnSalvarProcessoValorPagamento').click()");
                            Thread.Sleep(1500);
                            Waitforload(driver);

                            driver.Navigate().GoToUrl("https://mercadolivre.elaw.com.br/processoView.elaw");
                            ////*[@id='tabViewProcesso:pvp-dtProcessoValorResults']/div[2]/table/tbody/tr/td[3]/div - -xpath valor
                            pagamento.ID_ELaw = driver.FindElement(By.XPath("//*[@id='tabViewProcesso:pvp-dtProcessoValorResults']/div[2]/table/tbody/tr/td[3]/div")).Text;


                            GravarLog(pagamento, "INCLUSÃO BEM SUCEDIDA");
                            EscreveExcel novoItem = new EscreveExcel();
                            pagamento.Status = "INCLUSÃO BEM SUCEDIDA";
                            novoItem.GeraLogXLSX(dadosPagamento);

                            Thread.Sleep(1000);
                            //continue
                        }catch (Exception ex){

                            GravarLog(pagamento, $"ERRO: {ex.Message}");

                            EscreveExcel novoItem = new EscreveExcel();
                            pagamento.Status = ex.Message;
                            novoItem.GeraLogXLSX(dadosPagamento);
                            Waitforload(driver);
                        }
                    }
                    catch (Exception ex)
                    {
                        GravarLog(pagamento,$"ERRO: {ex.Message}");

                        EscreveExcel novoItem = new EscreveExcel();
                        //pagamento.Status = $"ERRO: Não foi possível seguir com a solicitação";
                        pagamento.Status = ex.Message;
                        novoItem.GeraLogXLSX(dadosPagamento);
                        Waitforload(driver);
                    }
                }
            }
        }

        private void GravarLogEncerrado(BaixarEntity processo, string message)
        {
            using (StreamWriter sw = new StreamWriter($@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\Log-Robo-Custas-ML-Processos-Encerrados.txt", true))
            {
                sw.WriteLine($"{processo.Processo} ; {DateTime.Now} ; {message} ");
            }
        }

        private void GravarLog(BaixarEntity pagamento, string message)
        {
            using (StreamWriter sw = new StreamWriter($@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\Log-Robo-Custas-ML.txt", true))
            {
                sw.WriteLine($"{pagamento.Processo} ; {DateTime.Now} ; {message} ");
            }
        }

        private void Waitforload(ChromeDriver driver)
        {
            Thread.Sleep(1000);
            IJavaScriptExecutor js = driver;
            int timeoutSec = 300;
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, timeoutSec));
            wait.Until(wd => js.ExecuteScript("return document.readyState").ToString() == "complete");
        }

        private static bool confereitem(ChromeDriver driver, string selector, string tipo = "xpath")
        {
            Thread.Sleep(500);

            try
            {
                switch (tipo)
                {
                    case "id":
                        if (driver.FindElement(By.Id(selector)).Displayed)
                        {
                            return false;
                        }
                        break;
                    case "class":
                        if (driver.FindElement(By.ClassName(selector)).Displayed)
                        {
                            return false;
                        }
                        break;
                    case "xpath":
                        if (driver.FindElement(By.XPath(selector)).Displayed && driver.FindElement(By.XPath(selector)).Enabled)
                        {
                            return false;
                        }
                        break;
                    case "link":
                        if (driver.FindElement(By.PartialLinkText(selector)).Displayed)
                        {
                            return false;
                        }
                        break;
                    default:
                        return true;
                        break;
                }


            }
            catch (Exception)
            {
                return true;

            }
            return true;
        }
        private bool ConfereitemClass(ChromeDriver driver, string xpath)
        {
            Thread.Sleep(1000);
            driver.SwitchTo().DefaultContent();
            try
            {
                if (!driver.FindElement(By.ClassName(xpath)).Displayed)
                {
                    Waitforload(driver);
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception)
            {
                return true;

            }
        }

        private bool ConfereitemId(ChromeDriver driver, string xpath)
        {
            Thread.Sleep(1000);
            driver.SwitchTo().DefaultContent();
            try
            {
                if (!driver.FindElement(By.Id(xpath)).Displayed)
                {
                    Waitforload(driver);
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception)
            {
                return true;

            }
        }

        private bool ConfereitemLinktext(ChromeDriver driver, String xpath)
        {
            Thread.Sleep(1000);
            driver.SwitchTo().DefaultContent();
            try
            {
                if (!driver.FindElement(By.LinkText(xpath)).Displayed)
                {
                    Waitforload(driver);
                    return true;
                }

                else
                {
                    return false;
                }

            }
            catch (Exception)
            {
                try
                {
                    if (!driver.FindElement(By.LinkText(xpath.Replace(Convert.ToChar(160), Convert.ToChar(32)).Replace(" ", ""))).Displayed)
                    {
                        Waitforload(driver);
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                }
                catch (Exception)
                {
                }
            }
            return true;
        }

        public String verificaDia()
        {
            DateTime diaAtual = DateTime.Today.AddDays(7);

            while (diaAtual.DayOfWeek != DayOfWeek.Monday && diaAtual.DayOfWeek != DayOfWeek.Thursday)
            {
                diaAtual = diaAtual.AddDays(1);
            }

            return diaAtual.ToString("dd/MM/yyyy");
        }

    }
}