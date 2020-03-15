using ConversorFaturasParaMinhasEconomias.Services;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ConversorFaturasParaMinhasEconomias
{
    class Program
    {
        static void Main(string[] args)
        {
            var caminhosDosArquivosOrigem = Directory.GetFiles(Directory.GetCurrentDirectory() + "\\Arquivos\\Origem\\");
            foreach (var localArquivoOrigem in caminhosDosArquivosOrigem)
            {
                string nomeDoBanco = string.Empty;
                IConversorFatura conversor;
                string dataTexto = string.Empty;
                DateTime data;
                using (StreamReader reader = new StreamReader(localArquivoOrigem))
                {
                    nomeDoBanco = reader.ReadLine();
                    conversor = ResolveConversor(nomeDoBanco);
                    dataTexto = reader.ReadLine();
                    data = DateTime.ParseExact(dataTexto, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                }

                var transacoes = conversor.Converter(localArquivoOrigem);

                var caminhoArquivoModelo = (Directory.GetCurrentDirectory() + "\\Arquivos\\Modelo_XLS.xls");
                var caminhoParaCopia = Directory.GetCurrentDirectory() + "\\Arquivos\\Destino";
                Directory.CreateDirectory(caminhoParaCopia);
                var caminhoArquivoCopia = $"{caminhoParaCopia}\\{data.ToString("yyyy-MM-dd")}-{nomeDoBanco}.xls";

                File.Copy(caminhoArquivoModelo, caminhoArquivoCopia, true);

                PreencherExcel(caminhoArquivoCopia, transacoes);

                var despesas = transacoes.Where(x => x.Valor < 0);
                var creditos = transacoes.Where(x => x.Valor > 0 && x.EhPagamentoFatura == false);
            }

            Console.ReadKey();
        }

        private static IConversorFatura ResolveConversor(string banco)
        {
            switch (banco.ToLowerInvariant())
            {
                case "santander":
                    return new ConversorFaturaSantander();
                default:
                    return new ConversorFaturaSantander();
            }
        }

        private static void PreencherExcel(string caminhoExcel, IList<Transacao> transacoes)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
                FileStream file = new FileStream(caminhoExcel, FileMode.OpenOrCreate);

                //Loads or open an existing workbook through Open method of IWorkbooks
                IWorkbook workbook = application.Workbooks.Open(file);
                IWorksheet worksheet = workbook.Worksheets[0];

                //var dataStyle = worksheet[$"A2"].CellStyle;
                //var valorStyle = worksheet[$"C2"].CellStyle;

                IStyle dataStyle = workbook.Styles.Add("DataStyle");
                dataStyle.NumberFormat = worksheet["A2"].CellStyle.NumberFormat;

                IStyle valorStyle = workbook.Styles.Add("ValorStyle");
                valorStyle.NumberFormat = worksheet["C2"].CellStyle.NumberFormat;

                var transacoesParaExcel = transacoes.Where(t => !t.EhPagamentoFatura).ToList();

                int quantidadeLinhasParaPular = 2;
                for (int i = 0; i < transacoesParaExcel.Count; i++)
                {
                    worksheet[$"A{i + quantidadeLinhasParaPular}"].Value = transacoesParaExcel[i].Data.ToString();
                    worksheet[$"A{i + quantidadeLinhasParaPular}"].CellStyle = dataStyle;

                    worksheet[$"B{i + quantidadeLinhasParaPular}"].Value = transacoesParaExcel[i].Descricao;

                    worksheet[$"C{i + quantidadeLinhasParaPular}"].Value = transacoesParaExcel[i].Valor.ToString();
                    worksheet[$"C{i + quantidadeLinhasParaPular}"].CellStyle = valorStyle;
                }

                //IStyle dataParcelasStyle = workbook.Styles.Add("DataParcelasStyle");
                //dataParcelasStyle.NumberFormat = worksheet["A2"].CellStyle.NumberFormat;
                //dataParcelasStyle.Font.Bold = true;

                //IStyle valorParcelaStyle = workbook.Styles.Add("ValorParcelaStyle");
                //valorParcelaStyle.NumberFormat = worksheet["C2"].CellStyle.NumberFormat;
                //valorParcelaStyle.Font.Bold = true;
                dataStyle.Font.Bold = true;
                valorStyle.Font.Bold = true;

                quantidadeLinhasParaPular = transacoesParaExcel.Count + 2;
                transacoesParaExcel = transacoesParaExcel.Where(t => t.EhParcelamento).SelectMany(t => t.ParcelamentoPai.TransacoesParceladas.Where(p => p.Parcela > t.Parcela)).ToList();
                for (int i = 0; i < transacoesParaExcel.Count(); i++)
                {
                    worksheet[$"A{i + quantidadeLinhasParaPular}"].Value = transacoesParaExcel[i].Data.ToString();
                    worksheet[$"A{i + quantidadeLinhasParaPular}"].CellStyle = dataStyle;
                    //worksheet[$"A{i + quantidadeLinhasParaPular}"].CellStyle.Font.Bold = true;

                    worksheet[$"B{i + quantidadeLinhasParaPular}"].Value = transacoesParaExcel[i].Descricao;
                    worksheet[$"B{i + quantidadeLinhasParaPular}"].CellStyle.Font.Bold = true;

                    worksheet[$"C{i + quantidadeLinhasParaPular}"].Value = transacoesParaExcel[i].Valor.ToString();
                    worksheet[$"C{i + quantidadeLinhasParaPular}"].CellStyle = valorStyle;
                    //worksheet[$"C{i + quantidadeLinhasParaPular}"].CellStyle.Font.Bold = true;
                }

                //Set the version of the workbook
                workbook.Version = ExcelVersion.Excel2016;

                file.Position = 0;
                workbook.SaveAs(file);

                workbook.Close();
            }

            //ExcelEngine excelEngine = new ExcelEngine();

            ////Instantiate the Excel application object
            //IApplication application = excelEngine.Excel;

            ////Assigns default application version
            //application.DefaultVersion = ExcelVersion.Excel2016;

            //FileStream file = new FileStream(caminhoExcel, FileMode.Open);

            //IWorkbook workbook = application.Workbooks.Open(file);

            ////Access first worksheet from the workbook.
            //IWorksheet worksheet = workbook.Worksheets[0];

            //for (int i = 0; i < transacoes.Count; i++)
            //{
            //    worksheet.Range[$"A{i+2}"].Text = transacoes[i].Data.ToString("dd/MM/yyyy");
            //    worksheet.Range[$"B{i + 2}"].Text = transacoes[i].Descricao;
            //    worksheet.Range[$"C{i + 2}"].Text = transacoes[i].Valor.ToString();
            //}

            ////Defining the ContentType for excel file.
            //string ContentType = "Application/msexcel";

            ////Define the file name.
            //string fileName = caminhoExcel;

            ////Creating stream object.
            //MemoryStream stream = new MemoryStream();

            ////Saving the workbook to stream in XLSX format
            //workbook.SaveAs(stream);

            //stream.Position = 0;

            ////Closing the workbook.
            //workbook.Close();

            ////Dispose the Excel engine
            //excelEngine.Dispose();

            //Creates a FileContentResult object by using the file contents, content type, and file name.
            //return File(stream, ContentType, fileName);
        }
    }
}
