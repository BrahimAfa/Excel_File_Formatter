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
       // List<BalanceFomr2> Balances2 = new List<BalanceFomr2>();

        private async void button1_Click(object sender, EventArgs e)
        {

            if (Directory.Exists(@"E:\BALANCE 2015 TEST.xlsx"))
            {
                MessageBox.Show(@"Plz Make a file in (E:\) Drive cuz there is no one!!!");
                return;

            }
            var wb = new XLWorkbook(@"E:\BALANCE 2015 TEST.xlsx");

            var ws = wb.Worksheet(1);
           
          //  var range = ws.RangeUsed();
         // var list = ;
        // BalanceFomr2.bal = await FillBalance2sasync(ws);
            dataGridView1.DataSource = fillbalance2List(ws.RangeUsed());

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
                BalanceFomr2.bal.Add(new BalanceFomr2()
               {
                   Compte = item.Cell(1).GetString(),
                   Intitule = item.Cell(2).GetString(),
                   Debit1 = debit1,
                   Credit1 = credit1,
                   Debit2 = debit2,
                   Credit2 = credit2
               });

           }
           return BalanceFomr2.bal;
       }

        //  Task<List<Balance>> FillBalancesasync(IXLWorksheet worksheet)
        //{

        //   return   Task.Run(() => fillbalanceList(worksheet.RangeUsed()));

        //}

        Task<List<BalanceFomr2>> FillBalance2sasync(IXLWorksheet worksheet)
          {

              return Task.Run(() => fillbalance2List(worksheet.RangeUsed()));

          }

        //List<Balance> dublicated = new List<Balance>();
        //List<Balance> dublicatedNew = new List<Balance>();
        List<BalanceFomr2> dublicatedNew2 = new List<BalanceFomr2>();
          List<DuplicatedList> list = new List<DuplicatedList>();
        private void button2_Click(object sender, EventArgs e)
        {
            #region GOld Commente
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

            //var query = from b in Balances2
            //            join l in list
            //            on b.Compte equals l.compte
            //            select b;
            //  Balances.Select(FirstOrDefault(c => c.Compte.Equals(c.Compte));
            #endregion
            var l = BalanceFomr2.bal;
             list = returnListOfDuplicated(l);

            //list.ForEach(x => anonym.Add(new DuplicatedList()
            //{
            //    Compte = x.compte,
            //    COUNT = x.count
            //}));

            //var dublicated = Balances2.Where(b => list.Any(x => x.compte == b.Compte)).OrderBy(w => w.Compte).ToList();

            //dataGridView2.DataSource = RemoveDuplicatedCompte(anonym).OrderBy(x => x.Compte).ToList();
            dataGridView2.DataSource = list.OrderBy(x=>x.Compte).ToList();
            label4.Text = dataGridView2.Rows.Count.ToString();


        }
        List<BalanceFomr2> RemoveDuplicatedCompte(List<DuplicatedList> Ls)
        {
            //Duplicated More Than 3 Times
            MessageBox.Show("In the Start of remove duplicate Func" + BalanceFomr2.bal.Count.ToString());

             var bal = BalanceFomr2.bal;
            List<BalanceFomr2> BalanceToRemove = new List<BalanceFomr2>();
            for (int j = 0; j < Ls.Count; j++)
            {
                for (int i = 0; i < bal.Count; i++)
            {
                
                    if (Ls[j].Compte == bal[i].Compte & Ls[j].COUNT == 3)
                    {

                        BalanceToRemove.Add(bal.FirstOrDefault(x => x.Compte == bal[i].Compte & x.Credit1 == 0 & x.Credit2 == 0 & x.Debit1 == 0 & x.Debit2 == 0));

                        //bal.Remove(bal.FirstOrDefault(x => x.Compte == bal[i].Compte & x.Credit1 == 0 & x.Credit2 == 0 & x.Debit1 == 0 & x.Debit2 == 0));
                        break;
                    }

                }
                    
                
            }
             BalanceToRemove.ForEach(x => bal.Remove(x));
            MessageBox.Show("In the ENd of remove duplicate Func"+BalanceFomr2.bal.Count.ToString());
            return bal.Where(b => returnListOfDuplicated(bal).Any(x => x.Compte == b.Compte)).OrderBy(w => w.Compte).ToList(); 
        }
        List<DuplicatedList> returnListOfDuplicated(List<BalanceFomr2> bal)
        {
            return bal.
                  GroupBy(x => x.Compte).
                  Where(b => b.Count() > 1).
                  Select(x => new DuplicatedList
                  {
                      Compte = x.Key,
                      COUNT = x.Count()
                  }).ToList();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            //dublicatedNew =  GetListBalClean();
            //dataGridView3.DataSource = dublicatedNew;
            //label6.Text = dataGridView3.Rows.Count.ToString();
            dublicatedNew2 =  RemoveDuplicatedCompte(list);
            dataGridView3.DataSource = dublicatedNew2.OrderBy(x => x.Compte).ToList() ;
            label6.Text = dataGridView3.Rows.Count.ToString();
        }

        //List<Balance> GetListBalClean()
        //{
        //    List<Balance> Bal = new List<Balance>();
        //    for (int i = 0; i < dataGridView2.Rows.Count; i++)
        //    {
        //        var item = dataGridView2.Rows[i];
        //        Bal.Add(new Balance()
        //        {
        //            Compte = item.Cells[0].Value.ToString(),
        //            Intitule = item.Cells[1].Value.ToString(),
        //            Debit = decimal.Parse(item.Cells[2].Value.ToString()),
        //            Credit = decimal.Parse(item.Cells[3].Value.ToString())
        //        });
        //    }
        //    for (int i = 0; i < Bal.Count; i++)
        //    {
        //        Bal[i].TotCredit = Bal[i].Credit + Bal[i + 1].Credit;


        //        Bal[i].TotDebit = Bal[i].Debit + Bal[i + 1].Debit;


        //        //dublicated[i].TotCredit = dublicated[i].Credit + dublicated[i + 1].Credit;
        //        //dublicated[i].TotDebit = dublicated[i].Debit + dublicated[i + 1].Debit;
        //        Bal.RemoveAt(i + 1);
        //    }
        //    return Bal;
        //}

        List<BalanceFomr2> GetListBal2Clean()
        {
            List<BalanceFomr2> Bal = dublicatedNew2;
            for (int i = 0; i < Bal.Count; i++)
            {

                Bal[i].TotCredit1 = Bal[i].Credit1 + Bal[i+1].Credit1;
                Bal[i].TotCredit2 = Bal[i].Credit2 + Bal[i +1].Credit2;
                Bal[i].TotDebit1 = Bal[i].Debit1 + Bal[i + 1].Debit1;
                Bal[i].TotDebit2 = Bal[i].Debit2 + Bal[i + 1].Debit2;

                Bal.RemoveAt(i+1);

            }
 



                //dublicated[i].TotCredit = dublicated[i].Credit + dublicated[i + 1].Credit;
                //dublicated[i].TotDebit = dublicated[i].Debit + dublicated[i + 1].Debit;
              
            
            return Bal;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //if (string.IsNullOrEmpty(textBox1.Text))
            //{
            //    dataGridView3.DataSource = dublicatedNew;
            //    return;
            //}

            //dataGridView3.DataSource = (dataGridView3.DataSource as List<Balance>).Where(x=> x.Compte.StartsWith(textBox1.Text)).ToList();
         
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                dataGridView3.DataSource = dublicatedNew2;
                return;
            }

            dataGridView3.DataSource = (dataGridView3.DataSource as List<BalanceFomr2>).Where(x => x.Compte.StartsWith(textBox1.Text)).ToList();


        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView4.DataSource = GetListBal2Clean();
            label8.Text = dataGridView4.Rows.Count.ToString();
            label9.Text = (dataGridView4.DataSource as List<BalanceFomr2>).Sum(x => x.TotCredit2).ToString();
            label10.Text = (dataGridView4.DataSource as List<BalanceFomr2>).Sum(x => x.TotDebit2).ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
                
        }
  
        private void button7_Click(object sender, EventArgs e)
        {
            MessageBox.Show(BalanceFomr2.bal.Count+"");
            var l = (dataGridView4.DataSource as List<BalanceFomr2>);
            dataGridView5.DataSource = BalanceFomr2.bal.Except(l).ToList();
            label12.Text = dataGridView5.Rows.Count.ToString();
        }
    }
}
