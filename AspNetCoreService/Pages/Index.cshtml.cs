using Microsoft.AspNetCore.Mvc.RazorPages;
using ConsoleApp1;

namespace AspNetCoreService.Pages
{
    public class IndexModel : PageModel
    {
        public void OnGet()
        {
            ConsoleApp1.Program.Main(null);
        }
    }
}
