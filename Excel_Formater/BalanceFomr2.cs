using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication4
{
    class BalanceFomr2
    {
        public BalanceFomr2(){}

        [DisplayName("N° Compte")]
        public string Compte { get; set; }


        [DisplayName("Intitule")]
        public string Intitule { get; set; }


        [DisplayName("Debit 1")]
        public decimal Debit1 { get; set; }



        [DisplayName("Credit 1")]
        public decimal Credit1 { get; set; }


        [DisplayName("Debit 2")]
        public decimal Debit2 { get; set; }


        [DisplayName("Credit 2")]
        public decimal Credit2 { get; set; }


        [DisplayName("TotDebit 1")]
        public decimal TotDebit1 { get; set; }


        [DisplayName("TotCredit 1")]
        public decimal TotCredit1 { get; set; }


        [DisplayName("TotDebit 2")]
        public decimal TotDebit2 { get; set; }


        [DisplayName("TotCredit 2")]
        public decimal TotCredit2 { get; set; }
    

    }
}
