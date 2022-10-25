using BLL.Entity;
using BLL.Robo;
using BLL.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MessageBox = System.Windows.Forms.MessageBox;

namespace Custas_MercadoLivre
{
    /// <summary>
    /// Interação lógica para MainWindow.xam
    /// </summary>
    ///   ML_login" value="hosana.costa"/> , ML_senha" value="Ailson2020@" />
    ///   pasta exemplo: 412686
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Execucao inicio = new Execucao();
            txtPagamento.Text = inicio.verificaDia();
        }

        private void Inicia_btn_Click(object sender, RoutedEventArgs e)
        {
            var data = Regex.Match(txtPagamento.Text, @"\d{2}\/\d{2}\/\d{4}").Value;

            if (data == "")
            {
                MessageBox.Show("Preencha a data corretamente!","Custas Mercado Livre");

                return;
            }

            System.Windows.Forms.OpenFileDialog arquivo = new System.Windows.Forms.OpenFileDialog();
            arquivo.ShowDialog();
            string caminho_do_arquivo = arquivo.FileName;

            if (string.IsNullOrEmpty(caminho_do_arquivo))
            {
                System.Windows.MessageBox.Show("Por favor Selecione um arquivo!", "Erro Arquivo");
                return;
            }

            ImportarExcel excel = new ImportarExcel();
            List<BaixarEntity> dados = excel.GetDadosPlanilha(caminho_do_arquivo);

            Execucao navegacao = new Execucao();
            navegacao.SolicitarPagamento(txtUsuario.Text, txtSenha.Password, txtPagamento.Text, dados);


        }

    }
}
