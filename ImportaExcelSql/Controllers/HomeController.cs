using ImportaExcelSql.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ImportaExcelSql.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            string filePath = string.Empty;
            try
            {
                if (file != null)
                {
                    string path = Server.MapPath("~/Uploads/");
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    filePath = path + Path.GetFileName(file.FileName);
                    string extension = Path.GetExtension(file.FileName);
                    file.SaveAs(filePath);

                    string conString = string.Empty;

                    switch (extension)
                    {
                        case ".xls": //Excel 97-03.
                            conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                            break;
                        case ".xlsx": //Excel 07 and above.
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=YES'";
                            break;
                    }

                    DataTable dt = new DataTable();
                    conString = string.Format(conString, filePath);

                    using (OleDbConnection connExcel = new OleDbConnection(conString))
                    {
                        using (OleDbCommand cmdExcel = new OleDbCommand())
                        {
                            using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                            {
                                cmdExcel.Connection = connExcel;

                                //Get the name of First Sheet.
                                connExcel.Open();
                                DataTable dtExcelSchema;
                                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                connExcel.Close();

                                //Read Data from First Sheet.
                                connExcel.Open();
                                cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                                odaExcel.SelectCommand = cmdExcel;
                                odaExcel.Fill(dt);
                                connExcel.Close();
                            }
                        }
                    }

                    conString = @"Server=NBDZAMBRANOB\SQLEXPRESS;Database=BaseDZ;Trusted_Connection=True;";
                    using (SqlConnection con = new SqlConnection(conString))
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            //Set the database table name.
                            sqlBulkCopy.DestinationTableName = "dbo.TB_IMP_IMPORTADATOS";

                            // Map the Excel columns with that of the database table, this is optional but good if you do
                            // 
                            sqlBulkCopy.ColumnMappings.Add("Id", "PK_IMP_IMPORTADATOS_ID");
                            sqlBulkCopy.ColumnMappings.Add("RutFacilitador", "TB_IMP_IMPORTADATOS_RutFacilitador");
                            sqlBulkCopy.ColumnMappings.Add("NombreFacilitador", "TB_IMP_IMPORTADATOS_NombreFacilitador");
                            sqlBulkCopy.ColumnMappings.Add("FechaIngreso", "TB_IMP_IMPORTADATOS_FechaIngreso");
                            sqlBulkCopy.ColumnMappings.Add("FechaAsignación", "TB_IMP_IMPORTADATOS_FechaAsignación");
                            sqlBulkCopy.ColumnMappings.Add("FechaEvaluación", "TB_IMP_IMPORTADATOS_FechaEvaluación");
                            sqlBulkCopy.ColumnMappings.Add("NombreEvaluador", "TB_IMP_IMPORTADATOS_NombreEvaluador");
                            sqlBulkCopy.ColumnMappings.Add("SectorModulo", "TB_IMP_IMPORTADATOS_SectorModulo");
                            sqlBulkCopy.ColumnMappings.Add("SubSectorModulo", "TB_IMP_IMPORTADATOS_SubSectorModulo");
                            sqlBulkCopy.ColumnMappings.Add("TipoModulo", "TB_IMP_IMPORTADATOS_TipoModulo");
                            sqlBulkCopy.ColumnMappings.Add("PlanFormativo", "TB_IMP_IMPORTADATOS_PlanFormativo");
                            sqlBulkCopy.ColumnMappings.Add("NombreModulo", "TB_IMP_IMPORTADATOS_NombreModulo");
                            sqlBulkCopy.ColumnMappings.Add("Estado", "TB_IMP_IMPORTADATOS_Estado");
                            sqlBulkCopy.ColumnMappings.Add("Correo", "TB_IMP_IMPORTADATOS_Correo");
                            sqlBulkCopy.ColumnMappings.Add("Teléfono", "TB_IMP_IMPORTADATOS_Teléfono");
                            sqlBulkCopy.ColumnMappings.Add("Comuna", "TB_IMP_IMPORTADATOS_Comuna");
                            sqlBulkCopy.ColumnMappings.Add("Región", "TB_IMP_IMPORTADATOS_Región");
                            sqlBulkCopy.ColumnMappings.Add("FechaEnvio", "TB_IMP_IMPORTADATOS_FechaEnvio");

                            con.Open();
                            sqlBulkCopy.WriteToServer(dt);
                            con.Close();
                        }
                    }
                }
                //if the code reach here means everthing goes fine and excel data is imported into database
                ViewBag.Success = "Archivo excel importado y guardado en la base de datos";

            }
            catch (Exception ex)
            {               
                ViewBag.Success = "Se genero un error" + ex.Message;
            }
            return View();
        }

        public ActionResult About()
        {
            List<TB_IMP_IMPORTADATOS> datos = new List<TB_IMP_IMPORTADATOS>();
            using (var context = new BaseDZEntities())
            {
                datos = context.TB_IMP_IMPORTADATOS.ToList();
            }
            ViewBag.Message = "Datos ingresados...";
            return View(datos);
        }

        public void ExportListUsingEPPlus()
        {
            List<TB_IMP_IMPORTADATOS> data = new List<TB_IMP_IMPORTADATOS>();
            using (var context = new BaseDZEntities())
            {
                data = context.TB_IMP_IMPORTADATOS.ToList();
            }

            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
            workSheet.Cells[1, 1].LoadFromCollection(data, true);
            using (var memoryStream = new MemoryStream())
            {
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //here i have set filname as Students.xlsx
                Response.AddHeader("content-disposition", "attachment;  filename=Importados.xlsx");
                excel.SaveAs(memoryStream);
                memoryStream.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.End();
            }
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}