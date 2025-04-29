using System;

namespace RadugaMassPrint.Models
{
    public class DocumentData
    {
        public string DocumentName { get; set; }
        public string Address { get; set; }
        public string BuildingName { get; set; }
        public string AccountName { get; set; }
        public int AgrmID { get; set; }
        public string AgreementNumber { get; set; }
        public string FileName { get; set; }
        public long DocID { get; set; }
        public DateTime OrderDate { get; set; }
        public decimal Sum { get; set; }
        public bool DifferentSum { get; set; } = false;
    }
}
