using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Pass : Form
    {
        public Pass()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (tbUser.Text == "" && tbPass.Text == "")
            {
                frmConfig fc = new frmConfig();
                fc.lbSaludo.Text = "Bienvenido DjRasek";
                this.Hide();
                fc.ShowDialog();
                this.Close();
            }
            else {
                MessageBox.Show(this, "Acceso Denegado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbPass.Text = "";
            }
        }
    }
}
