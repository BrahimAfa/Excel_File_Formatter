using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication4
{
    class Balance
    {
        public Balance()
        {

        }
        [DisplayName("N° Compte")]
        public string Compte { get; set; }
             [DisplayName("Intitule")]
        public string Intitule { get; set; }
             [DisplayName("Debit")]
        public decimal Debit { get; set; }
             [DisplayName("Credit")]
        public decimal Credit { get; set; }
             [DisplayName("TotCredit")]
        public decimal TotCredit { get; set; }
             [DisplayName("TotDebit")]
        public decimal TotDebit { get; set; }
    }
}
