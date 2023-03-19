using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OrdenadorWeb.Services;

namespace OrdenadorWeb.Pages;

public class Tempo : PageModel
{
    public PlanilhaService PlanilhaService { get; set; }
    public string TempoDeProcessamento { get; set; } = "0";

    public Tempo(PlanilhaService planilhaService)
    {
        PlanilhaService = planilhaService;
    }

    public IActionResult OnGet()
    {
        TempoDeProcessamento = PlanilhaService.TempoDeProcessamento.ToString(@"s\:fff");

        return Page();
    }

    public IActionResult OnPost()
    {
        return File(PlanilhaService.PlanilhaOrdenada, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"ExcelEmOrdem{PlanilhaService.TipoSelecionado}De{PlanilhaService.CriterioSelecionado}.xlsx");
    }
}
