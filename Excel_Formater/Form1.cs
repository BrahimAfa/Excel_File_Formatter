using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.IO;
//using _Excel = Microsoft.Office.Interop.Excel;
namespace WindowsFormsApplication4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //List<Balance> Balances = new List<Balance>();
        List<BalanceFomr2> Balances2 = new List<BalanceFomr2>();

        private async void button1_Click(object sender, EventArgs e)
        {
      
            
            var wb = new XLWorkbook(@"D:\BALANCE 2015 TEST.xlsx");

            var ws = wb.Worksheet(1);
           
          //  var range = ws.RangeUsed();
         // var list = ;
            dataGridView1.DataSource = await FillBalance2sasync(ws);

            label2.Text = dataGridView1.Rows.Count.ToString();





        }

       //private  List<Balance> fillbalanceList(IXLRange xL)
       // {

       //     xL.FirstRow().Delete();
       //     foreach (var item in xL.Rows())
       //     {
       //         var j = item.Cell(3).GetString();
       //         decimal debit = !string.IsNullOrEmpty( item.Cell(3).GetString())?decimal.Parse( item.Cell(3).GetString()):0;
       //         decimal credit = !string.IsNullOrEmpty( item.Cell(4).GetString())?decimal.Parse( item.Cell(4).GetString()) :0;
       //         Balances.Add(new Balance() 
       //         {
       //             Compte = item.Cell(1).GetString(),
       //             Intitule = item.Cell(2).GetString(),
       //             Debit = debit,
       //             Credit =credit
       //         });

       //     }
       //     return Balances;
       // }
        private List<BalanceFomr2> fillbalance2List(IXLRange xL)
       {

           xL.FirstRow().Delete();
           foreach (var item in xL.Rows())
           {
               var j = item.Cell(3).GetString();
               decimal debit1 = !string.IsNullOrEmpty(item.Cell(3).GetString()) ? decimal.Parse(item.Cell(3).GetString()) : 0;
               decimal credit1 = !string.IsNullOrEmpty(item.Cell(4).GetString()) ? decimal.Parse(item.Cell(4).GetString()) : 0;
               decimal debit2 = !string.IsNullOrEmpty(item.Cell(5).GetString()) ? decimal.Parse(item.Cell(5).GetString()) : 0;
               decimal credit2 = !string.IsNullOrEmpty(item.Cell(6).GetString()) ? decimal.Parse(item.Cell(6).GetString()) : 0;
               Balances2.Add(new BalanceFomr2()
               {
                   Compte = item.Cell(1).GetString(),
                   Intitule = item.Cell(2).GetString(),
                   Debit1 = debit1,
                   Credit1 = credit1,
                   Debit2 = debit2,
                   Credit2 = credit2
               });

           }
           return Balances2;
       }

        //  Task<List<Balance>> FillBalancesasync(IXLWorksheet worksheet)
        //{

        //   return   Task.Run(() => fillbalanceList(worksheet.RangeUsed()));

        //}

        Task<List<BalanceFomr2>> FillBalance2sasync(IXLWorksheet worksheet)
          {

              return Task.Run(() => fillbalance2List(worksheet.RangeUsed()));

          }

        List<Balance> dublicated = new List<Balance>();
        List<Balance> dublicatedNew = new List<Balance>();
        List<DuplicatedList> anonym = new List<DuplicatedList>();
        private void button2_Click(object sender, EventArgs e)
        {
            //var list = (from b in Balances
            //           where b.Compte == b.Compte
            //           select b) as List<Balance>;
            //Balances.Take(Balances.Count - 1).Where((index, item) => Balances[index + 1] == item);
            //var list = Balances.Where(
            //  b1 => Balances.Where(b2 => b1.Compte == b2.Compte).Count() != 0);
            //var list = Balances.
            //           GroupBy(x =>x.Compte).
            //           Where(b => b.Count() > 1).
            //           Select(x=>new {
            //               compte= x.Key
            //           }).ToList();
            //var dublicated = Balances.Where(b => list.Any(x => x.compte.Equals(b.Compte))).ToList();
            var list = Balances2.
                     GroupBy(x => x.Compte).
                     Where(b => b.Count() > 1).
                     Select(x => new
                     {
                         compte = x.Key,
                         count = x.Count()
                     }).ToList();
            foreach (var item in list)
            {
                anonym.Add(new DuplicatedList()
                {
                    Compte = item.compte,
                    COUNT = item.count
                });
            }
            //var query = from b in Balances2
            //            join l in list
            //            on b.Compte equals l.compte
            //            select b;
            var dublicated = Balances2.Where(b => list.Any(x => x.compte == b.Compte)).OrderBy(w=>w.Compte).ToList();

            dataGridView2.DataSource = list;
            label4.Text = dataGridView2.Rows.Count.ToString();



            //  Balances.Select(FirstOrDefault(c => c.Compte.Equals(c.Compte));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dublicatedNew =  GetListBalClean();
            dataGridView3.DataSource = dublicatedNew;
            label6.Text = dataGridView3.Rows.Count.ToString();
        }
       
        List<Balance> GetListBalClean()
        {
            List<Balance> Bal = new List<Balance>();
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                var item = dataGridView2.Rows[i];
                Bal.Add(new Balance()
                {
                    Compte = item.Cells[0].Value.ToString(),
                    Intitule = item.Cells[1].Value.ToString(),
                    Debit = decimal.Parse(item.Cells[2].Value.ToString()),
                    Credit = decimal.Parse(item.Cells[3].Value.ToString())
                });
            }
            for (int i = 0; i < Bal.Count; i++)
            {
                Bal[i].TotCredit = Bal[i].Credit + Bal[i + 1].Credit;


                Bal[i].TotDebit = Bal[i].Debit + Bal[i + 1].Debit;


                //dublicated[i].TotCredit = dublicated[i].Credit + dublicated[i + 1].Credit;
                //dublicated[i].TotDebit = dublicated[i].Debit + dublicated[i + 1].Debit;
                Bal.RemoveAt(i + 1);
            }
            return Bal;
        }
     
        //List<Balance> GetListBal2Clean()
        //{
        //    List<BalanceFomr2> Bal = new List<BalanceFomr2>();
        //    for (int i = 0; i < dataGridView2.Rows.Count; i++)
        //    {
        //        var item = dataGridView2.Rows[i];
        //        Bal.Add(new BalanceFomr2()
        //        {
        //            Compte = item.Cells[0].Value.ToString(),
        //            Intitule = item.Cells[1].Value.ToString(),
        //            Debit1 = decimal.Parse(item.Cells[2].Value.ToString()),
        //            Credit1 = decimal.Parse(item.Cells[3].Value.ToString()),
        //            Debit2 = decimal.Parse(item.Cells[4].Value.ToString()),
        //            Credit2 = decimal.Parse(item.Cells[5].Value.ToString()),

        //        });
        //    }
        //    for (int i = 0; i < Bal.Count; i++)
        //    {
        //                for (int j = 0; i <anonym.Count ; j++)
        //        {
        //            if (Bal[i].Compte == anonym[i].Compte)
        //            {
        //                for (int k = 0; k <  anonym[i].COUNT; k++)
        //                {

        //                    //Bal[i].TotCredit = Bal[i].Credit + Bal[i + 1].Credit;

        //                    //Bal[i].TotDebit = Bal[i].Debit + Bal[i + 1].Debit;
        //                }
        //            }
        //        }
             


        //        //dublicated[i].TotCredit = dublicated[i].Credit + dublicated[i + 1].Credit;
        //        //dublicated[i].TotDebit = dublicated[i].Debit + dublicated[i + 1].Debit;
        //        Bal.RemoveAt(i + 1);
        //    }
        //    return Bal;
        //}
        private void button4_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                dataGridView3.DataSource = dublicatedNew;
                return;
            }
            
            dataGridView3.DataSource = (dataGridView3.DataSource as List<Balance>).Where(x=> x.Compte.StartsWith(textBox1.Text)).ToList();


        }

        private void button6_Click(object sender, EventArgs e)
        {

        }
    }
}
