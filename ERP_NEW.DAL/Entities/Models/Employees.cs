using System;
using System.ComponentModel.DataAnnotations;

namespace ERP_NEW.DAL.Entities.Models
{
    public class Employees
    {
        [Key]
        public int EmployeeID { get; set; }
        public int AccountNumber { get; set; }
        public DateTime Engaged { get; set; }
        public DateTime Fired { get; set; }
    }
}
