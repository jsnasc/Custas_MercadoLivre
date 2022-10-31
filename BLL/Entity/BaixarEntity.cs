using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL.Entity
{
   public class BaixarEntity
    {
        public string Pasta_Cliente { get; set; }

        public string Valor { get; set; }

        public string Tipo_Pagamento{ get; set; } //Acordo, Condenação ou Custas

        public string Processo { get; set; }

        public string Proc_Jud { get; set; } //JEC ou Vara Cível

        public string Caminho_Arquivo { get; set; }

        public string Pasta_GED { get; set; }

        public string Nome_Arquivo { get; set; }

        public string Status { get; set; }

        public string diaHora { get; set; }

        public string ID_ELaw { get; set; }
    }
}
