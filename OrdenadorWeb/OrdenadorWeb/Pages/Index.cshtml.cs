using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OrdenadorWeb.Services;

namespace OrdenadorWeb.Pages;

public class Index : PageModel
{
    public PlanilhaService PlanilhaService { get; set; }

    public Index(PlanilhaService planilhaService)
    {
        PlanilhaService = planilhaService;
    }

    public IActionResult OnPost(IFormFile arquivoExcel, string criterioDeOrdenacao, string tipoDeOrdenacao)
    {
        if (MostrarMensagensDeErro(arquivoExcel, criterioDeOrdenacao, tipoDeOrdenacao))
            return Page();

        MemoryStream planilhaDesordenada = new MemoryStream();
        arquivoExcel.CopyToAsync(planilhaDesordenada);

        PlanilhaService.ProcessarPlanilha(planilhaDesordenada, criterioDeOrdenacao, tipoDeOrdenacao);

        return RedirectToPage("Tempo");
    }

    public bool MostrarMensagensDeErro(IFormFile arquivoExcel, string criterioSelecionado, string tipoSelecionado)
    {
        if (arquivoExcel == null || arquivoExcel.Length == 0)
        {
            ViewData["MensagemDeErroExcel"] = "Nenhum arquivo EXCEL selecionado.";
            return true;
        }
        else if (!arquivoExcel.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            ViewData["MensagemDeErroExcel"] = "Tipo de arquivo não suportado.";
            return true;
        }

        if (string.IsNullOrEmpty(criterioSelecionado))
        {
            ViewData["MensagemDeErroCriterio"] = "Escolha pelo menos um critério de ordenação.";
            return true;
        }

        if (string.IsNullOrEmpty(tipoSelecionado))
        {
            ViewData["MensagemDeErroTipo"] = "Escolha pelo menos um tipo de ordenação.";
            return true;
        }

        return false;
    }
}