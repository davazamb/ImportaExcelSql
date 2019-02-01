using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ImportaExcelSql.Models
{
    public class ImportaDatosView
    {
        public int IdImporta { get; set; }
        public string RutFacilitador { get; set; }
        public string NombreFacilitador { get; set; }
        public string FechaIngreso { get; set; }
        public string FechaAsignación { get; set; }
        public string FechaEvaluación { get; set; }
        public string NombreEvaluador { get; set; }
        public string SectorModulo { get; set; }
        public string SubSectorModulo { get; set; }
        public string TipoModulo { get; set; }
        public string PlanFormativo { get; set; }
        public string NombreModulo { get; set; }
        public string Estado { get; set; }
        public string Correo { get; set; }
        public string Teléfono { get; set; }
        public string Comuna { get; set; }
        public string Región { get; set; }
        public string FechaEnvio { get; set; }
    }
}