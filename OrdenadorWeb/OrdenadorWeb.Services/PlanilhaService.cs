using OfficeOpenXml;
using OfficeOpenXml.Style;
using OrdenadorWeb.Models;
using OrdenadorWeb.ServiceContracts;
using OrdenadorWeb.ServiceContracts.Enums;
using System.Drawing;

namespace OrdenadorWeb.Services;

public class PlanilhaService : IPlanilhaService
{
    public Planilha Planilha { get; set; }
    public TimeSpan TempoDeProcessamento { get; set; }
    public string CriterioSelecionado { get; set; }
    public string TipoSelecionado { get; set; }
    public MemoryStream PlanilhaOrdenada { get; set; } = new();

    public PlanilhaService()
    {
        Planilha = new Planilha();
        CriterioSelecionado = string.Empty;
        TipoSelecionado = string.Empty;
    }

    public void ProcessarPlanilha(MemoryStream planilhaDesordenada, string criterioDeOrdenacao, string tipoDeOrdenacao)
    {
        DateTime TempoInicial = DateTime.Now;

        ConverterExcelParaPlanilha(planilhaDesordenada);

        CriterioSelecionado = criterioDeOrdenacao;
        TipoSelecionado = tipoDeOrdenacao;

        Enum.TryParse(criterioDeOrdenacao, out CriterioDeOrdenacao criterioSelecionado);
        Enum.TryParse(tipoDeOrdenacao, out TipoDeOrdenacao tipoSelecionado);

        OrdenarPlanilha(criterioSelecionado, tipoSelecionado);

        ConverterPlanilhaParaExcel();

        TempoDeProcessamento = DateTime.Now - TempoInicial;
    }

    public void ConverterExcelParaPlanilha(MemoryStream arquivoExcel)
    {
        arquivoExcel.Position = 0;

        using (ExcelPackage package = new ExcelPackage(arquivoExcel))
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelWorksheet planilha = package.Workbook.Worksheets.First();
            int numeroDeLinha = planilha.Dimension.End.Row;

            for (int i = 2; i <= numeroDeLinha; i++)
            {
                string nome = planilha.Cells[i, 1].Value.ToString();
                int idade = int.Parse(planilha.Cells[i, 2].Value.ToString());
                Pessoa pessoa = new Pessoa(nome, idade);

                Planilha.Dados.Add(pessoa);
            }
        }
    }

    public void ConverterPlanilhaParaExcel()
    {
        PlanilhaOrdenada = new();

        using (ExcelPackage package = new ExcelPackage())
        {
            ExcelWorksheet planilha = package.Workbook.Worksheets.Add("ListaOrdenada");
            int index = 2;

            planilha.Cells[1, 1].Value = "Nome";
            planilha.Cells[1, 2].Value = "Idade";

            ExcelRange celulas = planilha.Cells["A1:B1"];
            celulas.Style.Font.Bold = true;
            celulas.Style.Fill.PatternType = ExcelFillStyle.Solid;
            celulas.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            celulas.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            foreach (Pessoa pessoa in Planilha.Dados)
            {
                planilha.Cells[index, 1].Value = pessoa.Nome;
                planilha.Cells[index, 2].Value = pessoa.Idade;
                index++;
            }

            planilha.Column(1).AutoFit();
            planilha.Column(2).AutoFit();

            package.SaveAs(PlanilhaOrdenada);

            PlanilhaOrdenada.Position = 0;
        }
    }

    public void OrdenarPlanilha(CriterioDeOrdenacao criterioDeOrdenacao, TipoDeOrdenacao tipoDeOrdenacao)
    {
        if (tipoDeOrdenacao == TipoDeOrdenacao.Crescente)
        {
            if (criterioDeOrdenacao == CriterioDeOrdenacao.Nome)
            {
                Planilha.Dados = Planilha.Dados.OrderBy(x => x.Nome).ThenBy(x => x.Idade).ToList();
            }
            else
            {
                Planilha.Dados = Planilha.Dados.OrderBy(x => x.Idade).ThenBy(x => x.Nome).ToList();
            }
        }
        else
        {
            if (criterioDeOrdenacao == CriterioDeOrdenacao.Nome)
            {
                Planilha.Dados = Planilha.Dados.OrderByDescending(x => x.Nome).ThenByDescending(x => x.Idade).ToList();
            }
            else
            {
                Planilha.Dados = Planilha.Dados.OrderByDescending(x => x.Idade).ThenByDescending(x => x.Nome).ToList();
            }
        }
    }
}
