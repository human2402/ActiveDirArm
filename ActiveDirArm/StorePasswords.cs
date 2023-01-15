using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using IronXL;

namespace ActiveDirArm
{
    internal class StorePasswords
    {
        string fileYear = "666";
        string fileDate = "666";
        public StorePasswords() {
            DateTime lastWriteTime = File.GetLastWriteTime($@"{Directory.GetCurrentDirectory()}\userlist_new.csv");
            var fileMonth = lastWriteTime.Month;
            string fileMonthStr="";
            if (fileMonth < 10) { 
                fileMonthStr = "0" + Convert.ToString(fileMonth);
            }

            this.fileYear = Convert.ToString(lastWriteTime.Year);
            this.fileDate = lastWriteTime.Day + "_" + fileMonthStr;
            Console.WriteLine(fileYear + "    " + fileDate);
        }
        public bool DoesExistInFolder()
        {
            return File.Exists($@"{Directory.GetCurrentDirectory()}\CSVs\{fileYear}\{fileDate}.csv");
        }
        public void StoreInFoler()
        {
            Console.WriteLine($@"{Directory.GetCurrentDirectory()}\userlist_new.csv"+"           "+ $@"{Directory.GetCurrentDirectory()}\CSVs\{fileYear}\{fileDate}.csv");
            if (!DoesExistInFolder())
            {
                //create a folder im CSVs corresponding to the year of modification of the file
                if (!Directory.Exists($@"{Directory.GetCurrentDirectory()}\CSVs\{fileYear}")) {
                    Directory.CreateDirectory($@"{Directory.GetCurrentDirectory()}\CSVs\{fileYear}");
                }

                File.Move($@"{Directory.GetCurrentDirectory()}\userlist_new.csv", $@"{Directory.GetCurrentDirectory()}\CSVs\{fileYear}\{fileDate}.csv");
                MessageBox.Show("Файл перенесен!", "Ура");

            } else {
                Random rnd = new Random();
                int new_index = rnd.Next(1, 100);
                MessageBox.Show("В папке " + $@"{Directory.GetCurrentDirectory()}\CSVs\{fileYear}\" + " уже присутствует файл " + $"{fileDate}.csv. Копируем копию {new_index}.", "Ну как же так-тоо?"     ) ;
                File.Move($@"{Directory.GetCurrentDirectory()}\userlist_new.csv", $@"{Directory.GetCurrentDirectory()}\CSVs\{fileYear}\{fileDate}-{new_index}.csv");
            }
        }




        public void StoreInExcel ()
        {
            WorkBook wb = WorkBook.Load("localBook.xlsx");
            WorkSheet wsh = wb.WorkSheets.First();
            bool FoundFlag = false;
            foreach(WorkSheet sheet in wb.WorkSheets)
            {
                if (sheet.Name.Equals(fileYear))
                {
                    Console.WriteLine("Yeah found it" + sheet.Name);
                    wsh= sheet;
                    FoundFlag = true; break;
                }
            }
            if ( !FoundFlag ) {
                wsh = wb.CreateWorkSheet(fileYear);
            }

            int startLine = 0;
            for (int i = 1; i<99999; i++)
            {
                //Console.WriteLine(wsh["A" + i].Int32Value);
                if (wsh["A" + i].StringValue+":0" == ":0")
                {
                    Console.WriteLine(i);
                    startLine = i;break;
                }
            }
            Console.WriteLine(startLine);
            wsh["A" + startLine].Value = this.fileDate.ToString();
            wsh["A" + startLine].Style.BackgroundColor="#8bfca9";


            WorkBook userList = WorkBook.Load("userlist_new.csv");
            WorkSheet wsh_userList = userList.WorkSheets.First();

            for (int i = 1; i<200; i++)
            {
                Console.WriteLine(wsh_userList["A" + i].Value);
                wsh["A" + Convert.ToString(i + startLine)].Value = wsh_userList["A" + i].Value;
                wsh["B" + Convert.ToString(i + startLine)].Value = wsh_userList["B" + i].Value;
                wsh["C" + Convert.ToString(i + startLine)].Value = wsh_userList["C" + i].Value;
                wsh["D" + Convert.ToString(i + startLine)].Value = wsh_userList["D" + i].Value;
                wsh["E" + Convert.ToString(i + startLine)].Value = wsh_userList["E" + i].Value;
                wsh["F" + Convert.ToString(i + startLine)].Value = wsh_userList["F" + i].Value;

                //Console.WriteLine(wsh_userList["A" + i].StringValue);    
                if (wsh_userList["A" + i].StringValue =="") { break; }
            }


            wb.Save();

            string[] syncLocation = File.ReadAllLines($@"{Directory.GetCurrentDirectory()}/defaultExcelSyncLocation.txt");
            

            DialogResult shouldSync = MessageBox.Show("Пароли сохранены в localBook.xlsx. Хотите сохранить в " + syncLocation[0] + " ?", "Nice", MessageBoxButtons.YesNo);
            if (shouldSync == DialogResult.Yes) {
                try
                {
                    wb.SaveAs(syncLocation[0]);
                } catch(Exception ex)
                {
                    MessageBox.Show(ex.Message + " Вот так не получилось", "Ну ничего себе как так (");
                }
            }
            //try
            //{
            //    wsh = wb.GetWorkSheet(this.fileYear);
            //} catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //        Console.WriteLine("yeah");
            //}
        }
    }
}
