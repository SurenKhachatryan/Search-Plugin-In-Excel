using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFindMatchRows.Models
{
    public class ResultModel
    {
        public string TableName { get; set; }
        public List<Excel.Range> Rows { get; set; }
    }
}
