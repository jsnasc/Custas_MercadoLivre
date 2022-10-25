using BLL.Entity;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL.Services
{
   public class ImportarExcel
    {
        public List<BaixarEntity> GetDadosPlanilha(string caminho)
        {
            List<BaixarEntity> dados = new List<BaixarEntity>();

            try
            {
                using (XLWorkbook workbook = new XLWorkbook(caminho))
                {
                    int linha = 9;

                    while (!workbook.Worksheet(1).Cell(linha, 1).Value.ToString().Contains("Valor Não processado"))
                    {
                        BaixarEntity data = new BaixarEntity();

                        data.Pasta_Cliente = Convert.ToString(workbook.Worksheet(1).Cell(linha, 7).Value);
                        data.Valor = string.Format("{0:F2}",workbook.Worksheet(1).Cell(linha, 10).Value);
                        data.Tipo_Pagamento = Convert.ToString(workbook.Worksheet(1).Cell(linha, 11).Value);
                        data.Processo = Convert.ToString(workbook.Worksheet(1).Cell(linha, 17).Value);
                        data.Pasta_GED = Convert.ToString(workbook.Worksheet(1).Cell(linha, 16).Value);
                        data.Nome_Arquivo = Convert.ToString(workbook.Worksheet(1).Cell(linha, 21).Value);
                        data.Pasta_GED = data.Pasta_GED.Replace("IRIS00", "");
                        data.Caminho_Arquivo = $@"\\192.168.0.30\GED\" + data.Pasta_GED + @"\" + data.Nome_Arquivo;

                        dados.Add(data);
                        linha++;
                    }
                }
            }
            catch (Exception) { }
            return dados;
        }
    }
}
