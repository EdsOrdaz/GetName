using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Management;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GetName
{
    public partial class Form1 : Form
    {
        #region SQL
        private String select = "SELECT ltrim(rtrim(nombre)) as nombre, ltrim(rtrim(apepat)) as paterno, ltrim(rtrim(apemat)) as materno,email FROM [Nom2001].[dbo].[nomtrab] where status='A' ORDER BY Nombre asc";
        private String conexionsql = "server=40.76.105.1,5055; database=Nom2001;User ID=reportesUNNE;Password=8rt=h!RdP9gVy; integrated security = false ; MultipleActiveResultSets=True";
        private String conexionsqllast = "server=148.223.153.37,5314; database=InfEq;User ID=eordazs;Password=Corpame*2013; integrated security = false ; MultipleActiveResultSets=True";
        
        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        public static List<String[]> lista = new List<String[]>();


        public static RichTextBox firma;

        public static string PrimeraMayuscula(String value)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value);
        }

        private void loadimagen()
        {
            pictureBox1.Enabled = true;
            pictureBox1.BringToFront();
            pictureBox1.Visible = true;
        }

        private void hideimagen()
        {
            pictureBox1.Visible = false;
            pictureBox1.Enabled = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if(backgroundWorker1.IsBusy != true)
            {
                loadimagen();
                lista.Clear();
                backgroundWorker1.RunWorkerAsync();
            }
        }
        
        private void TextBox1_KeyUp(object sender, KeyEventArgs e)
        {
            dataGridView1.Rows.Clear();
            if(!string.IsNullOrEmpty(textBox1.Text))
            {
                foreach (String[] empleado in lista)
                {
                    if (empleado[5].Contains(textBox1.Text.ToUpper()))
                    {
                        dataGridView1.Rows.Add(empleado[3], empleado[4], empleado[6]);
                    }
                }
            }
        }

        /*
        private String generar()
        {
            int longitudContrasenia;

            string mayus = (mayusculas.Checked == true) ? "ABCDEFGHIJKLMNOPQRSTUVWXYZ" : "";
            string nums = (numeros.Checked == true) ? "1234567890" : "";
            string punts = (puntuacion.Checked == true) ? "#$%&()='?¿¡!+*}{[]-_:;" : "";

            longitudContrasenia = Convert.ToInt32(ncaracteres.Value);

            Random rdn = new Random();
            string caracteres = "abcdefghijklmnopqrstuvwxyz"+mayus+nums+punts;
            int longitud = caracteres.Length;
            char letra;
            string contraseniaAleatoria = string.Empty;
            for (int i = 0; i < longitudContrasenia; i++)
            {
                letra = caracteres[rdn.Next(longitud)];
                contraseniaAleatoria += letra.ToString();
            }
            textBox2.Text = contraseniaAleatoria;
            portapapeles(contraseniaAleatoria);
            return contraseniaAleatoria;
        }
        */

        private void portapapeles(String texto)
        {
            if (!String.IsNullOrEmpty(texto))
            {
                Clipboard.SetText(texto);
                label1.ForeColor = System.Drawing.Color.Green;
                label1.Text = "Portapapeles: " + texto;
            }
        }
        
        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            /*
            if (e.ColumnIndex == this.dataGridView1.Columns["name"].Index && e.RowIndex != -1)
            {
                portapapeles(dataGridView1.Rows[e.RowIndex].Cells["name"].Value.ToString());
            }
            if (e.ColumnIndex == this.dataGridView1.Columns["id"].Index && e.RowIndex != -1)
            {
                portapapeles(dataGridView1.Rows[e.RowIndex].Cells["id"].Value.ToString());
            }
            if (e.ColumnIndex == this.dataGridView1.Columns["email"].Index && e.RowIndex != -1)
            {
                portapapeles(dataGridView1.Rows[e.RowIndex].Cells["email"].Value.ToString());
            }
            */
        }
        
        private void guardardatos()
        {
            try
            {
                using (SqlConnection conexion2 = new SqlConnection(conexionsqllast))
                {
                    String nombre = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                    String correo = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    String usuario = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    Console.WriteLine(correo);


                    if (String.IsNullOrEmpty(textBox2.Text))
                    {
                        MessageBox.Show("Debes ingresar una contraseña.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    String pass = VerPass.Encriptar(textBox2.Text);
                    String insert = "INSERT INTO GetName VALUES ('"+nombre+"','"+usuario+"','"+correo+"', '"+pass+"', '"+DateTime.Now+"')";
                    Console.WriteLine(insert);
                    conexion2.Open();
                    SqlCommand comm2 = new SqlCommand(insert, conexion2);
                    comm2.ExecuteReader();
                    MessageBox.Show("Registro Guardado.", "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Hide();

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error al guardar datos.\n\nMensaje: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void Button1_Click(object sender, EventArgs e)
        {
            //generar();
            guardardatos();
        }

        private void BackgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                using (SqlConnection conexion = new SqlConnection(conexionsql))
                {
                    conexion.Open();
                    SqlCommand comm = new SqlCommand(select, conexion);
                    SqlDataReader nwReader = comm.ExecuteReader();
                    while (nwReader.Read())
                    {
                        String[] n = new String[7];
                        n[0] = PrimeraMayuscula(nwReader["nombre"].ToString().ToLower());
                        n[1] = PrimeraMayuscula(nwReader["paterno"].ToString().ToLower());
                        n[2] = (string.IsNullOrEmpty(nwReader["materno"].ToString())) ? "" : PrimeraMayuscula(nwReader["materno"].ToString().ToLower());
                        n[3] = n[0] + " " + n[1] + " " + n[2];
                        char nom = n[0][0];
                        String id;
                        if (string.IsNullOrEmpty(nwReader["materno"].ToString()))
                        {
                            id = nom.ToString() + n[1];
                        }
                        else
                        {
                            char a = n[2][0];
                            id = nom.ToString() + n[1] + a.ToString();
                        }
                        n[4] = id.ToLower();
                        n[5] = n[3].ToUpper();
                        n[6] = nwReader["email"].ToString();
                        lista.Add(n);
                    }
                }
            }
            catch (System.Exception ex)
            {
                e.Cancel = true;
                MessageBox.Show("Error en la busqueda\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if(e.Cancelled)
            {
                Hide();
            }
            hideimagen();
        }
    }
}

