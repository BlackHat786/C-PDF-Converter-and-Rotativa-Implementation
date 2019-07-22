using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace PDFDownloadsTest1.Models
{
    public class Employee
    {
        [Key]
        public int EmployeeID { get; set; }

        [Display(Name ="Employee Name")]
        public string EmpName { get; set; }

        [Display(Name ="Employee Surname")]
        public string EmpSName { get; set; }

        public string Gender { get; set; }

        public string City { get; set; }

        [Display(Name ="Date Hired")]
        public DateTime DateHired { get; set; }

        
    }
}