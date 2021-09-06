using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERP_NEW.DAL.Entities.QueryModels
{
    public class EmployeePhoto
    {
        [Key]
        public int EmployeeID { get; set; }
        public byte[] Photo { get; set; }
    }
}
