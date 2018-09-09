using System.Web;
using System.IO;
using System.Web.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using reg = System.Text.RegularExpressions;
using System.Linq;

namespace CodingCraft.AulaRegex.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        
        public ActionResult AnaliseExcel() {
            HttpPostedFileBase file = Request.Files["FileUpload"];

            string dados = "";

            using (var excel = new ExcelPackage(file.InputStream)) {
                
                ExcelWorksheet worksheet = excel.Workbook.Worksheets[1];
                
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;
                
                for (int row = start.Row + 1; row <= end.Row; row++) {
                    var rowData = "";
                    for (int col = start.Column; col < end.Column; col++)
                    {
                        rowData += worksheet.Cells[1, col].Value + " " + worksheet.Cells[row, col].Value + " ";
                    }
                    dados += rowData;
                }
            }

            #region Verificação de cidades
            var cidades = new Dictionary<String, int>();
            var regexCidade = new reg.Regex(@"CIDADE (?<cid>[\w\.\s]+) UF", reg.RegexOptions.IgnoreCase);

            foreach (reg.Match ocorrencia in regexCidade.Matches(dados))
            {
                var cid = ocorrencia.Groups["cid"].Value;

                if (!cidades.ContainsKey(cid))
                {
                    cidades[cid] = 1;
                    continue;
                }
                cidades[cid]++;
            }

            List<SelectListItem> topCidades = new List<SelectListItem>();
            
            foreach (var item in cidades.OrderByDescending(d => d.Value).Take(5).ToList())
            {
                topCidades.Add(new SelectListItem
                {
                    Text = item.Key,
                    Value = item.Value.ToString()
                });
            }
            #endregion

            #region Verificação de Períodos
            var periodos = new Dictionary<String, int>();
            var regexPeriodo = new reg.Regex(@"PERIDOOCORRENCIA (?<periodo>[\w\.\s]+) DATACOMUNICACAO", reg.RegexOptions.IgnoreCase);

            foreach (reg.Match ocorrencia in regexPeriodo.Matches(dados))
            {
                var per = ocorrencia.Groups["periodo"].Value;

                if (!periodos.ContainsKey(per))
                {
                    periodos[per] = 1;
                    continue;
                }
                periodos[per]++;
            }


            List<SelectListItem> topPeriodos = new List<SelectListItem>();

            foreach (var item in periodos.OrderByDescending(d => d.Value).Take(5).ToList())
            {
                topPeriodos.Add(new SelectListItem
                {
                    Text = item.Key,
                    Value = item.Value.ToString()
                });
            }
            #endregion

            #region Verificação de Logradouros
            var logradouros = new Dictionary<String, int>();
            var regexLogradouros = new reg.Regex(@"LOGRADOURO (?<logradouro>[\w\.\s]+) UF", reg.RegexOptions.IgnoreCase);

            foreach (reg.Match ocorrencia in regexLogradouros.Matches(dados))
            {
                var log = reg.Regex.Replace(ocorrencia.Groups["logradouro"].Value, @"NUMERO (?<logr>[\w\.\s]+) CIDADE", " - ");

                if (log != "  -  ")
                {
                    if (!logradouros.ContainsKey(log))
                    {
                        logradouros[log] = 1;
                        continue;
                    }
                    logradouros[log]++;
                }
            }
            
            List<SelectListItem> topLogradouros = new List<SelectListItem>();

            foreach (var item in logradouros.OrderByDescending(d => d.Value).Take(5).ToList())
            {
                topLogradouros.Add(new SelectListItem
                {
                    Text = item.Key,
                    Value = item.Value.ToString()
                });
            }
            #endregion

            ViewBag.Periodos = topPeriodos;
            ViewBag.Cidades = topCidades;
            ViewBag.Logradouros = topLogradouros;

            return PartialView("_AnaliseExcel");
        }
    }
}