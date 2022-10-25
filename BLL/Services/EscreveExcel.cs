using BLL.Entity;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL.Services
{
    public class EscreveExcel
    {
        public void GeraLogXLSX (List<BaixarEntity> dados)
        {
            //$@"{Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)}\Desktop\ML-Processos-Encerrados.xlsx"
            try
            {
                using (var wbook = new XLWorkbook())
                {
                    var ws = wbook.Worksheets.Add("Status-Solicitações");
                    wbook.SaveAs($@"{Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)}\Desktop\Relatorio-Custas-ML" + DateTime.Now.ToString("-dd-MM-yyyy") + ".xlsx");

                    ws.Cell("A1").Value = "Pasta Cliente";
                    ws.Cell("B1").Value = "Processo";
                    ws.Cell("C1").Value = "Status";
                    ws.Cell("D1").Value = "Data e Hora";
                    ws.Range("A1:D1").Style.Font.Bold = true;


                    int l = 2;

                    foreach (BaixarEntity item in dados)
                    {
                        ws.Cell($"A{l}").Value = item.Pasta_Cliente;
                        ws.Cell($"B{l}").Value = item.Processo;
                        ws.Cell($"C{l}").Value = item.Status;
                        ws.Cell($"E{l}").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
                        l++;
                    }

                    wbook.Save();

                    /*IXLWorksheet ws = workbook.Worksheets.Add("Status-Encerrado");
                    ws.Cell(1,1).Value = "Pasta_Cliente";
                    ws.Cell(1,1).Style.Font.Bold = true;

                    ws.Cell(1,2).Value = "Processo";
                    ws.Cell(1,2).Style.Font.Bold = true;

                    ws.Cell(1, 3).Value = "Status";
                    ws.Cell(1, 3).Style.Font.Bold = true;

                    ws.Cell(linha,1).Value = dado.Pasta_Cliente;
                    ws.Cell(linha,2).Value = dado.Processo;
                    ws.Cell(linha,3).Value = "ENCERRADO";*/

                    //workbook.SaveAs($@"{Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)}\Desktop\ML-Processos-Encerrados.xlsx");
                }
            }
            catch (Exception ex)
            { 
            
            }
        }
    }


}
