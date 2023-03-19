using OrdenadorWeb.Models;
using OrdenadorWeb.ServiceContracts.Enums;

namespace OrdenadorWeb.ServiceContracts;

public interface IPlanilhaService
{
    public Planilha Planilha { get; set; }
    public TimeSpan TempoDeProcessamento { get; set; }
    public string CriterioSelecionado { get; set; }
    public string TipoSelecionado { get; set; }
    public MemoryStream PlanilhaOrdenada { get; set; }

    public void ProcessarPlanilha(MemoryStream planilhaDesordenada, string criterioDeOrdenacao, string tipoDeOrdenacao);

    public void OrdenarPlanilha(CriterioDeOrdenacao criterioDeOrdenacao, TipoDeOrdenacao tipoDeOrdenacao);

    public void ConverterExcelParaPlanilha(MemoryStream arquivoExcel);

    public void ConverterPlanilhaParaExcel();
}
