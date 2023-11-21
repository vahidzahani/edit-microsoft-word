using System;
using System.IO;
using System.Windows.Forms;
using Word=Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Threading;
using System.Linq;

namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {
        

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {



        }
        private string fn_getpath()
        {
            string filePath = Process.GetCurrentProcess().MainModule.FileName;
            int index = filePath.LastIndexOf("\\");
            if (index >= 0)
            {
                filePath = filePath.Substring(0, index);
            }
            return(filePath);
        }
      
        public string getPersianNumber(string data)

        {

            for (int i = 48; i < 58; i++)
            {
                data = data.Replace(Convert.ToChar(i), Convert.ToChar(1728 + i));
            }
         return data;


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void fn_create_report(string fname)
        {
            //MessageBox.Show(fname);
            //return;
            Word.Application word = new Word.Application();
            word.Visible = true;
            word.WindowState = Word.WdWindowState.wdWindowStateNormal;
            Word.Document doc = word.Documents.Open(fn_getpath() + @"\gozaresh.docx");

            string filePath = fn_getpath() + @"\"+fname+".txt";
            
            StreamReader reader = new StreamReader(filePath);
            string a = "";
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                if (a == "")
                {
                    a = line;
                }
                else
                {

                    string vahid = getPersianNumber(line); //culture.NumberFormat.NumberToString(1944, NumberStyles.Any, culture);
                    vahid = FN_reverser(vahid);
                    doc.Range().Find.Execute(FindText: a, ReplaceWith: vahid, Replace: Word.WdReplace.wdReplaceAll);
                    a = "";
                }
            }
            reader.Close();
            File.Delete(fn_getpath() + @"\" + fname + ".txt");


            doc.SaveAs2(fn_getpath() + @"\"+fname+".docx");
            doc.Close();
            word.Quit();

        }
        private void fn_get_list_config()
        {
            string path = Directory.GetCurrentDirectory();

            // دریافت لیست فایل های موجود در مسیر برنامه
            string[] files = Directory.GetFiles(path);

            // چاپ نام فایل های با پسوند txt
            listBox1.Items.Clear();
            foreach (string file in files)
            {
                if (Path.GetExtension(file) == ".txt")
                {
                    //Console.WriteLine(file);
                    string fname = Path.GetFileNameWithoutExtension(file);
                    if ( fname== "config")continue;
                    Thread.Sleep(3000);
                    fn_create_report(fname);
                    listBox1.Items.Add(file);
                }
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            fn_get_list_config();
            timer1.Enabled = true;

        }
        private string FN_reverser(string str)
        {
            //string str = "1944/04/08/01/05/02";
            if (str.IndexOf('/') != -1)
            {
                string[] parts = str.Split('/');
                Array.Reverse(parts);
                string reversedDate = string.Join("/", parts);
                return (reversedDate);
            }
            else { 
                return str;
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            
            
            
        }
    }
}
