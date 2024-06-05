using Excel_Modification_Task.Helpers;
using Excel_Modification_Task.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace Excel_Modification_Task.Controllers
{
	public class HomeController : Controller
	{
		private readonly ILogger<HomeController> _logger;
		private readonly IWebHostEnvironment _webHostEnvironment;

		public HomeController(ILogger<HomeController> logger, IWebHostEnvironment webHostEnvironment)
		{
			_logger = logger;
			this._webHostEnvironment = webHostEnvironment;
		}

		public IActionResult Index()
		{
			return View();
		}


		[HttpPost]
		public async Task<IActionResult> ProcessFile(IFormFile file)
		{
			if (file is null || file.Length == 0 || !Path.GetExtension(file.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
			{
				ViewBag.Error = "Please Upload Valid Excel File";
				return View(nameof(Index));
			}
			var uploadsfolder = Path.Combine(_webHostEnvironment.WebRootPath, "Uploads");
			if (!Directory.Exists(uploadsfolder))
			{
				Directory.CreateDirectory(uploadsfolder);
			}
			var filePath = Path.Combine(uploadsfolder, file.FileName);
			using (var fileStream = new FileStream(filePath, FileMode.Create))
			{
				await file.CopyToAsync(fileStream);
			}
			var fileData = DocumentSetting.ProcessFile(filePath);
			var relativeFilePath= "/Uploads/"+file.FileName;
			var filePath1 = Url.Content(relativeFilePath);
			ViewBag.FilePath = filePath1;
			return View(nameof(Result), fileData);
		}


		public IActionResult Result()
		{
			return View();
		}

		[ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
		public IActionResult Error()
		{
			return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
		}
	}
}
