using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OriginGradeTrans
{
    public partial class Form1 : Form
    {
        static TextBox debugTB;

        static void Out(string InfoString)
        {
            debugTB.AppendText(InfoString + "\r\n");
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            debugTB = textBox1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            if (!File.Exists(openFileDialog1.FileName))
                return;
            Translator TransT = new Translator();
            TransT.SetDebugOut(Out);
            TransT.GradesTableTranslate(openFileDialog1.FileName);
        }
    }
}
