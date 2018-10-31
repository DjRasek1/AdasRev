using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApplication1
{   
    public partial class frmConfig : Form
    {
        public frmConfig()
        {
            InitializeComponent();

            if (File.Exists("Config.txt"))
            {
                string[] siglas = new string[5];
                string[] rutas = new string[5];

                Form1.extConfig(siglas, rutas); 

                tbRutas1.Text = rutas[0];
                tbRutas2.Text = rutas[1];
                tbRutas3.Text = rutas[2];
                tbRutas4.Text = rutas[3];
                tbRutas5.Text = rutas[4];
                tbSiglas1.Text = siglas[0];
                tbSiglas2.Text = siglas[1];
                tbSiglas3.Text = siglas[2];
                tbSiglas4.Text = siglas[3];
                tbSiglas5.Text = siglas[4];
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRuta1_Click(object sender, EventArgs e)
        {
            btnRutasClick(tbRutas1);
        }

        private void btnRuta2_Click(object sender, EventArgs e)
        {
            btnRutasClick(tbRutas2);
        }

        private void btnRuta3_Click(object sender, EventArgs e)
        {
            btnRutasClick(tbRutas3);
        }

        private void btnRuta4_Click(object sender, EventArgs e)
        {
            btnRutasClick(tbRutas4);
        }

        private void btnRuta5_Click(object sender, EventArgs e)
        {
            btnRutasClick(tbRutas5);
        }

        private void btnRutasClick(TextBox tb)
        {
            tb.Clear();
            fbDialog.ShowDialog();
            tb.Paste(fbDialog.SelectedPath.ToString());
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            string[] rutas = { tbRutas1.Text, tbRutas2.Text, tbRutas3.Text, tbRutas4.Text, tbRutas5.Text };
            string[] siglas = { tbSiglas1.Text, tbSiglas2.Text, tbSiglas3.Text, tbSiglas4.Text, tbSiglas5.Text };
            
            if (File.Exists("config.txt"))
            {
                File.Delete("config.txt");
            }

            File.Create("config.txt").Close();

            using (StreamWriter sw = new StreamWriter("config.txt"))
            {
                for (int i = 0; i <= 4; i++)
                {
                    sw.WriteLine(siglas[i] + "," + rutas[i]);
                }
            }
            MessageBox.Show("Cierre el programa que los cambios tengan efecto");
            this.Close();
            
        }
    }
}
