using System.Collections.Generic;

namespace ConversorFaturasParaMinhasEconomias.Services
{
    public class Parcelamento
    {
        public int Parcelas { get; set; }

        public List<Transacao> TransacoesParceladas { get; set; } = new List<Transacao>();
    }
}
