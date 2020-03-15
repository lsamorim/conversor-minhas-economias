using System.Collections.Generic;

namespace ConversorFaturasParaMinhasEconomias.Services
{
    public interface IConversorFatura
    {
        IList<Transacao> Converter(string localArquivo);
    }
}
