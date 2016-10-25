using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.IO;
using ClosedXML.Excel;

namespace MVCCLOSEDXML.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/
        public ActionResult Index()
        {
            return View();
        }
        public void Excel()
        {
            /*http://closedxml.codeplex.com/documentation
             * Documentación oficial de la libreria para poder manipular el excel como gustes.
             * */
            DataTable dt = new DataTable("GridView_Data");
            //Columnas
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("NOMBRE", typeof(string));
            dt.Columns.Add("EMAIL", typeof(string));
            dt.Columns.Add("ESTADO", typeof(string));
            //Datos
            dt.Rows.Add(1, "FERNANDO BACHUR", "XXXXXXXXX@GMAIL.COM", "ACTIVO");
            dt.Rows.Add(2, "MICHAEL LOBOS", "XXXXXXXXX@GMAIL.COM", "ACTIVO");
            dt.Rows.Add(3, "ALEJANDRO COMAS", "XXXXXXXXX@GMAIL.COM", "ACTIVO");
            dt.Rows.Add(4, "ROBERTO LAZO", "XXXXXXXXX@GMAIL.COM", "ACTIVO");
            dt.Rows.Add(5, "VICTOR RAMOS", "XXXXXXXXX@GMAIL.COM", "ACTIVO");
            //Creacion y entrega FBM
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);

                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=Listado.xlsx");
                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }
        }
	}
}