using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ConversorFaturasParaMinhasEconomias.Services
{
    public class ConversorFaturaSantander : IConversorFatura
    {
        public IList<Transacao> Converter(string localArquivo)
        {
            using (StreamReader reader = new StreamReader(localArquivo))
            {
                var template = reader.ReadLine();
                var dataTexto = reader.ReadLine();
                var data = DateTime.ParseExact(dataTexto, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                List<Transacao> transacoes = new List<Transacao>();
                while (reader.Peek() >= 0)
                {
                    var transacao = ConverterLinha(data, reader.ReadLine());
                    if (transacoes == null) continue;

                    transacoes.Add(transacao);
                }

                return transacoes;
            }
        }

        public Transacao ConverterLinha(DateTime data, string linha)
        {
            var colunaUs = linha.IndexOf("US$");
            var colunaRS = linha.IndexOf("R$");

            var descricao = linha.Substring(11, colunaUs - 12);
            var valorUSTexto = linha.Substring(colunaUs, colunaRS - colunaUs);
            var valorRSTexto = linha.Substring(colunaRS);
            var valorUS = decimal.Parse(valorUSTexto.Substring(4), CultureInfo.GetCultureInfo("pt-BR"));
            var valorRS = decimal.Parse(valorRSTexto.Substring(3), CultureInfo.GetCultureInfo("pt-BR"));

            var transacao = new Transacao();
            transacao.Data = data;
            transacao.Descricao = descricao;
            transacao.Valor = valorRS * -1;
            transacao.EhPagamentoFatura = (descricao == "PAGAMENTO DE FATURA");

            string pattern = @"\(\d\d\/\d\d\)";
            Match m = Regex.Match(linha, pattern, RegexOptions.IgnoreCase);
            if (m.Success)
            {
                var parcelaTexto = m.Value.Substring(1, 2);
                var maximoParcelasTexto = m.Value.Substring(4, 2);

                transacao.EhParcelamento = true;
                transacao.Parcela = int.Parse(parcelaTexto);
                transacao.MaximoParcela = int.Parse(maximoParcelasTexto);

                var parcelamento = new Parcelamento()
                {
                    Parcelas = transacao.MaximoParcela
                };
                transacao.ParcelamentoPai = parcelamento;

                var transacoesParceladas = new List<Transacao>() { transacao };

                int mesesOffset = 1;
                for (int i = transacao.Parcela + 1; i <= transacao.MaximoParcela; i++)
                {
                    transacoesParceladas.Add(new Transacao
                    {
                        ParcelamentoPai = parcelamento,
                        Data = transacao.Data.AddMonths(mesesOffset),
                        Descricao = transacao.Descricao.Replace(m.Value, $"({i.ToString("00")}/{transacao.MaximoParcela.ToString("00")})"),
                        Valor = transacao.Valor,
                        EhPagamentoFatura = transacao.EhPagamentoFatura,
                        EhParcelamento = true,
                        MaximoParcela = transacao.MaximoParcela,
                        Parcela = i
                    });

                    mesesOffset++;
                }

                parcelamento.TransacoesParceladas = transacoesParceladas;
            }

            return transacao;
        }
    }
}
