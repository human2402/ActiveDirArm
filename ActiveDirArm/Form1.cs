using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace ActiveDirArm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            void CheckFileLocation()
            {
                //check for correct app location NO NEED NO MORE
                //bool doesFileExist = File.Exists(@"C:\ADARM\ActiveDirArm.exe");
                //if (!doesFileExist) {
                //    MessageBox.Show("Необходимо расположить exe файл в папке ADARM на диске С! Мне просто оченьь сложно было делать всякие динамические адрессации, так-что как-то так", "Сис админ к успеху шел");
                //    Application.Exit();
                //}
                //check for userList

                bool doesUserFileExist = File.Exists($@"{Directory.GetCurrentDirectory()}\userlist.csv");
                if (!doesUserFileExist)
                {
                    DialogResult shouldContinue = MessageBox.Show("UserList.scv не найден. Продолжить?", "Чего-то не хватает", MessageBoxButtons.YesNo);
                    if (shouldContinue == DialogResult.No) { Application.Exit(); }
                }
                //check for userList with passwords
                bool doesPasswordsExist = File.Exists($@"{Directory.GetCurrentDirectory()}\userlist_new.csv");
                if (doesPasswordsExist)
                {
                    DialogResult shouldContinue = MessageBox.Show("UserList_New.scv найден. Возможно, пароли из него не были записаны в бд. Записать?", "Что-то лишнее", MessageBoxButtons.YesNo);
                    if (shouldContinue == DialogResult.Yes)
                    {
                        StorePasswords sp = new StorePasswords();
                        sp.StoreInExcel();
                        sp.StoreInFoler();
                    }
                }
            }
                Console.WriteLine(Directory.GetCurrentDirectory());
                CheckFileLocation();
                
            }

        private void Form1_Load(object sender, EventArgs e)
        {      
            
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProcessStartInfo ps = new ProcessStartInfo();
            ps.FileName = "cmd.exe";
            ps.WindowStyle = ProcessWindowStyle.Normal;
            ps.Arguments = $@"/k cscript.exe {Directory.GetCurrentDirectory()}/NewPass.vbs /";
            Process.Start(ps);
            //System.Console.ReadKey();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ProcessStartInfo ps = new ProcessStartInfo();
            ps.FileName = "cmd.exe";
            ps.WindowStyle = ProcessWindowStyle.Normal;
            ps.Arguments = $@"/k cscript.exe {Directory.GetCurrentDirectory()}/NewUser.vbs /";
            Process.Start(ps);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ProcessStartInfo ps = new ProcessStartInfo();
            ps.FileName = "cmd.exe";
            ps.WindowStyle = ProcessWindowStyle.Normal;
            ps.Arguments = $@"/k cscript.exe {Directory.GetCurrentDirectory()}/NewHome.vbs /";
            Process.Start(ps);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            StorePasswords sp = new StorePasswords();
            sp.StoreInExcel();
            sp.StoreInFoler();
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }
    }
}
