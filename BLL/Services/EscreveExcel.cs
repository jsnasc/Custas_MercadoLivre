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
                    ws.Cell("C1").Value = "Valor";
                    ws.Cell("D1").Value = "Status";
                    ws.Cell("E1").Value = "Data";
                    ws.Cell("F1").Value = "ID_Custa_Solicitada";
                    ws.Range("A1:F1").Style.Font.Bold = true;


                    int l = 2;

                    foreach (BaixarEntity item in dados)
                    {
                        ws.Cell($"A{l}").Value = item.Pasta_Cliente;
                        ws.Cell($"B{l}").Value = item.Processo;
                        ws.Cell($"C{l}").Value = item.Valor;
                        ws.Cell($"D{l}").Value = item.Status;
                        ws.Cell($"E{l}").Value = DateTime.Now.ToString("dd/MM/yyyy");
                        ws.Cell($"F{l}").SetValue(item.ID_ELaw);
                        l++;
                    }

                    wbook.Save();
                    //workbook.SaveAs($@"{Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)}\Desktop\ML-Processos-Encerrados.xlsx");
                }
            }
            catch (Exception ex)
            { 
            
            }
        }
    }


}
