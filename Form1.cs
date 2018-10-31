using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        #region Recupera las siglas y rutas de las estaciones
            string[] siglas = new string[5];
            string[] rutas = new string[5];

            extConfig(siglas, rutas);

            this.cbSelec.Items.AddRange(siglas);
            this.cbSelec.Text = siglas[0];

        #endregion

        }

        //Boton Revisar
        private void btRev_Click(object sender, EventArgs e)
        {
            string archivo = "";
            string file = "";
            string ruta = "";
            string rutaAudio = "";
            DateTime inicio = fInic.Value;
            DateTime fin = fFin.Value;
            string[] siglas = new string[5];
            string[] rutas = new string[5];

            extConfig(siglas, rutas);

            if (!dataGridView1.RowCount.Equals(0))
            {
                dataGridView1.Rows.Clear();
            }

            ruta = relacRutas(ruta, rutas);

            if (fInic.Value > fFin.Value)   
            {
                MessageBox.Show("La fecha inical no puede ser mayor a la fecha final");
            }
            else
            {
                while (inicio.DayOfYear <= fin.DayOfYear)
                {
                    file = FormFecha(inicio);
                    archivo = ruta + @"asciilist\" + cbSelec.SelectedItem + "_" + file + ".txt";
                    rutaAudio = ruta + @"AUDIO\ADASCOM\";


                    if (cbSelec.Visible == true)
                    {
                        progressBar1.Step = 50;
                        progressBar1.PerformStep();
                        lbProgreso.Text = "Leyendo " + cbSelec.SelectedItem.ToString() + " del dia " + inicio.ToShortDateString();
                        lbProgreso.Refresh();
                        protocRev(archivo, rutaAudio, inicio, file);
                    }

                    #region rbTodas está seleccionado
                    else
                    {
                        for (int i = 0; i < cbSelec.Items.Count; i++)
                        {
                            cbSelec.SelectedIndex = i;
                            ruta = relacRutas(ruta, rutas);
                            progressBar1.Step = 20;
                            progressBar1.PerformStep();
                            lbProgreso.Text = "Leyendo " + cbSelec.SelectedItem.ToString() + " del dia " + inicio.ToShortDateString();
                            lbProgreso.Refresh();
                            archivo = ruta + @"asciilist\" + cbSelec.SelectedItem + "_" + file + ".txt";
                            rutaAudio = ruta + @"AUDIO\ADASCOM\";
                            protocRev(archivo, rutaAudio, inicio, file);
                        }
                    }
                    #endregion

                    inicio = inicio.AddDays(1);
                    progressBar1.Value = 0;
                    lbProgreso.Text = "";
                }
                MessageBox.Show("Busqueda Terminada");
            }
        }

        //Relaciona las rutas con las siglas
        private string relacRutas(string ruta, string[] rutas)
        {
            switch (cbSelec.SelectedIndex)
            {
                case 0: ruta = rutas[0];
                    break;

                case 1: ruta = rutas[1];
                    break;

                case 2: ruta = rutas[2];
                    break;

                case 3: ruta = rutas[3];
                    break;

                case 4: ruta = rutas[4];
                    break;
            }
            return ruta;
        }

        //Protocolo de revisión.
        private void protocRev(string archivo, string rutaAudio, DateTime inicio, string file)
        {
            string x;
            string y = "";
            string audio = "00";
            string hora = "";
            int carpAudio = 00;
            bool existe = false;
            string[] codigo = { "", "" };

            if (File.Exists(archivo))
            {
                using (StreamReader leer = new StreamReader(archivo))
                {
                    while (!leer.EndOfStream)
                    {
                        x = leer.ReadLine();

                        if (x.Length > 10 && x.Substring(0, 2) == "#C")
                        {
                            hora = obtieneHora(x); 
                        }
                        codigo = separaCodigo(x);

                        string directorio = rutaAudio + audio + @"\";

                        if (y != "#C F" && codigo[0] != "")
                        {
                            while (Directory.Exists(directorio))
                            {
                                string buscar = directorio + codigo[0] + ".wav";

                                if (File.Exists(buscar))
                                {
                                    existe = true;
                                    audio = "00";
                                    carpAudio = 0;
                                    break;
                                }
                                else { existe = false; }

                                carpAudio++;
                                audio = carpAudio.ToString();

                                if (audio.Length == 1)
                                {
                                    audio = "0" + audio;
                                }
                                directorio = rutaAudio + audio + @"\";
                            }
                            if (existe == false)
                            {
                                escribeTabla(inicio, hora, codigo);
                                audio = "00";
                                carpAudio = 0;
                            }
                        }
                    }
                }
            }
            else { MessageBox.Show("El archivo " + cbSelec.SelectedItem.ToString() + "_" + file + ".txt no existe", "Advertencia"); }
        }

        //Boton Radio todas
        private void estTodas_CheckedChanged(object sender, EventArgs e)
        {
            cbSelec.Visible = false;
        }

        //Boton Radio Seleccionar
        private void estSelec_CheckedChanged(object sender, EventArgs e)
        {
            cbSelec.Visible = true;
        }

        //Boton Listado a Excel
        private void btList_Click(object sender, EventArgs e)
        {
            var excelApp = new excel.Application();
            excel.Workbooks book = excelApp.Workbooks;
            book.Add(true);
            excel.Range rango = excelApp.get_Range("A1", "E1");

            int indiceColumna = 0;
            int indiceFila = 0;
            foreach (DataGridViewColumn col in dataGridView1.Columns)  //Columnas
            {
                indiceColumna++;
                excelApp.Cells[1, indiceColumna] = col.Name;
            }

            indiceFila = 0;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                indiceFila++;
                indiceColumna = 0;
                foreach (DataGridViewColumn col in dataGridView1.Columns)  //Filas
                {
                    indiceColumna++;
                    excelApp.Cells[indiceFila + 1, indiceColumna] = row.Cells[col.Name].Value;
                    string fin = "E" + indiceFila.ToString();
                    rango = excelApp.Range["A1", fin];
                }
            }

            rango.Font.Name = "Verdana";
            excelApp.get_Range("A1", "E1").Font.Bold = true;
            excelApp.get_Range("A1", "E1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            rango.Borders.LineStyle = excel.XlLineStyle.xlContinuous;
            rango.Columns.AutoFit();

            excelApp.Visible = true;
        }

        //Bloque que escribe en la tabla sin repetir codigos.
        private void escribeTabla(DateTime fecha, string hora, string[] codigo)
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[3].Value.ToString() == codigo[0] && dataGridView1.Rows[i].Cells[0].Value.ToString() == cbSelec.SelectedItem.ToString())// && dataGridView1.Rows[i].Cells[1].Value.ToString() == fecha.ToShortDateString())
                {
                    codigo[0] = "";
                    break;
                }
            }

            if (codigo[0] != "")
            {
                dataGridView1.Rows.Add(cbSelec.SelectedItem, fecha.ToShortDateString(), hora, codigo[0], codigo[1]);
            }
        }

        //Bloque que separa el codigo del contrato.
        private static string[] separaCodigo(string x)
        {
            string[] codigo = { "", "" };

            if (x.Length > 15 && x.Length < 22 && x.Substring(0, 1) != "L" && x.Substring(0, 1) != "#")
            {
                codigo = x.Split(',');
            }
            return codigo;
        }

        //Bloque que obiente la hora del corte.
        private static string obtieneHora(string x)
        {
            string hora = "";
            string y;
                y = x.Substring(0, 4);
                if (y == "#C C")
                {
                    hora = x.Substring(16, 8);
                }
            return hora;
        }

        //Bloque que da formato a la fecha YYYYMMDD
        public string FormFecha(DateTime fecha)
        {
            string year = fecha.Year.ToString();
            string month = fecha.Month.ToString();
            string day = fecha.Day.ToString();

            if (month.Length != 2)
            {
                month = "0" + month;
            }

            if (day.Length != 2)
            {
                day = "0" + day;
            }

            string date = year + month + day;

            return date;
        }

        //Lanza el módulo de informacion
        private void acercaDeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 ab = new AboutBox1();
            ab.ShowDialog();
        }

        //Lanza el módulo de configuración
        private void configuracionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Pass pass = new Pass();
            pass.Show();
        }

        //Extrae la información de configuración
        public static void extConfig(string[] siglas, string[] rutas)
        {
            string[] lineas = File.ReadAllLines("config.txt");
            string[] ambos = { };
            int indice = 0;

            foreach (string item in lineas)
            {
                //for (int i = 0; i <= lineas.Length - 1; i++)
                //{
                ambos = item.Split(',');
                siglas[indice] = ambos[0];
                rutas[indice] = ambos[1];
                indice++;
                //} 
            }
        }
    }
}