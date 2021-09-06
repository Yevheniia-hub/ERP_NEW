using System;

namespace ERP_NEW.BLL.DTO.SelectedDTO
{
    public class EmployeesInfoDTO
    {
        public int EmployeeID { get; set; }
        public decimal AccountNumber { get; set; }
        public string Fio { get; set; }
        public string FullName { get; set; }
        public int? ProfessionID { get; set; }
        public int? DepartmentID { get; set; }
        public string ProfessionName { get; set; }
        public string DepartmentName { get; set; }
        public byte[] UserPhoto { get; set; }
        public DateTime DateBegin { get; set; }
        public DateTime DateEnd { get; set; }
        //public int? SupplierId { get; set; }

        public string Checked { get; set; }
        public bool Selected { get; set; }
    }
}
