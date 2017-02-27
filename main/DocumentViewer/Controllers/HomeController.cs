using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.Mvc;
using DocumentParser.src.parser;
using DocumentViewer.Models;
using Newtonsoft.Json;

namespace DocumentViewer.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/
        public ActionResult Index()
        {
            return View(new DocumentViewModel
                {
                    Objective = new List<string>(),
                    Achievement = new List<string>(),
                    Certification = new List<string>(),
                    ExperienceSummary = new List<string>(),
                    Education = new List<string>(),
                    ProjectExperience = new List<string>(),
                    Skill = new List<string>(),
                    Handle = new List<string>()
                });
        }

        // This action handles the form POST and the upload
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                var fileName = Path.GetFileName(file.FileName);

                if (fileName != null)
                {
                    var extension = Path.GetExtension(fileName);
                    const string mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                    if (string.Equals(extension, ".docx") && file.ContentType.Equals(mimeType))
                    {

                        var appDataPath = Server.MapPath("~/App_Data");
                        var path = Path.Combine(appDataPath, "Upload", fileName);
                        file.SaveAs(path);

                        var dictionaryPath = Path.Combine(appDataPath, "KeywordDictionary.xml");

                        var modelData = Engine.Parse(path, dictionaryPath);

                        return View(JsonConvert.DeserializeObject<DocumentViewModel>(modelData));
                    }
                }
            }

            return View();
        }
    }
}
