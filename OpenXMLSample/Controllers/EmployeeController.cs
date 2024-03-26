using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OpenXMLSample.Models;
using System.Security.AccessControl;
using System.Xml.Linq;
using System.IO;

namespace OpenXMLSample.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EmployeeController : ControllerBase
    {
        [HttpPost]
        public async Task<IActionResult> ImportEmployee()
        {
            var files = Request.Form.Files;
            if (!files.Any())
            {
                return BadRequest();
            }
            var file = files.FirstOrDefault();
            string fileName = await Helper.FileHelper.UpLoad(file);
            string path = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "uploads", fileName);
            var word = WordprocessingDocument.Open(path, true);
            var paragraphs = word.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();
            var employee = new Employee()
            {
                userId = paragraphs[EmployeeIndex.userId].InnerText.GetEmployeeProp(),
                name = paragraphs[EmployeeIndex.name].InnerText.GetEmployeeProp(),
                objective = paragraphs[EmployeeIndex.objective].InnerText.GetEmployeeProp(),
                educations = paragraphs[EmployeeIndex.educations].InnerText.GetEmployeeProp(),
                certifications = paragraphs[EmployeeIndex.certifications].InnerText.GetEmployeeProp(),
                hobbies = paragraphs[EmployeeIndex.hobbies].InnerText.GetEmployeeProp(),
                skills = paragraphs[EmployeeIndex.skills].InnerText.GetEmployeeProp(),
                avatarFileId = long.Parse(paragraphs[EmployeeIndex.avatarFileId].InnerText.GetEmployeeProp()),
                languageId = long.Parse(paragraphs[EmployeeIndex.languageId].InnerText.GetEmployeeProp()),
                isActive = bool.Parse(paragraphs[EmployeeIndex.isActive].InnerText.GetEmployeeProp()),
                hightLight = paragraphs[EmployeeIndex.hightLight].InnerText.GetEmployeeProp(),
                strength = paragraphs[EmployeeIndex.strength].InnerText.GetEmployeeProp(),
                urlGithub = paragraphs[EmployeeIndex.urlGithub].InnerText.GetEmployeeProp(),
                presenter = paragraphs[EmployeeIndex.presenter].InnerText.GetEmployeeProp(),
                honorsAndAwards = paragraphs[EmployeeIndex.honorsAndAwards].InnerText.GetEmployeeProp(),
                position = paragraphs[EmployeeIndex.position].InnerText.GetEmployeeProp()
            };
            var exps = word.MainDocumentPart.RootElement.Descendants<DocumentFormat.OpenXml.Drawing.Table>().ToList();
            return Ok(employee);
        }
    }
}
