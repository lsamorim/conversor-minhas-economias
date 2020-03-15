using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConversorFaturasParaMinhasEconomias.Services
{
    public class Transacao
    {
        public DateTime Data { get; set; }

        public string Descricao { get; set; }

        public decimal Valor { get; set; }

        public bool EhPagamentoFatura { get; set; }

        public bool EhParcelamento { get; set; }

        public Parcelamento ParcelamentoPai { get; set; }

        public int Parcela { get; set; }

        public int MaximoParcela { get; set; }
    }
}
