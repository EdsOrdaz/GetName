using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GetName
{
    public partial class Editar : Form
    {
        public static String buscar = "";
        public Editar()
        {
            InitializeComponent();
        }

        private void Editar_Load(object sender, EventArgs e)
        {
            buscar = "";
            nombre.Text = VerPass.edit_nombre;
            usuario.Text = VerPass.edit_usuario;
            correo.Text = VerPass.edit_correo;
            pass.Text = VerPass.edit_pass;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(String.IsNullOrEmpty(nombre.Text.Trim()) || String.IsNullOrEmpty(usuario.Text.Trim()) || String.IsNullOrEmpty(correo.Text.Trim()) || String.IsNullOrEmpty(pass.Text.Trim()))
            {
                MessageBox.Show("No debe haber campos vacios", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                using (SqlConnection conexion2 = new SqlConnection(VerPass.conexionsqllast))
                {
                    String pass_nuevo = VerPass.Encriptar(pass.Text);
                    conexion2.Open();
                    String sql2 = "UPDATE GetName SET nombre = '" + nombre.Text + "', usuario = '" + usuario.Text + "', correo = '" + correo.Text + "', password = '" + pass_nuevo + "', fecha = '"+DateTime.Now+"' WHERE id = '" + VerPass.edit_ID + "'";
                    SqlCommand comm2 = new SqlCommand(sql2, conexion2);
                    SqlDataReader nwReader2 = comm2.ExecuteReader();
                }
                MessageBox.Show("Registro editado con exito.", "Editar", MessageBoxButtons.OK, MessageBoxIcon.Information);
                buscar = usuario.Text;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al editar información.\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult msj = MessageBox.Show("¿Quieres eliminar este registro?", "Eliminar", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if(msj == DialogResult.Yes)
            {
                try
                {
                    using (SqlConnection conexion2 = new SqlConnection(VerPass.conexionsqllast))
                    {
                        conexion2.Open();
                        String sql2 = "DELETE FROM GetName WHERE id = '" + VerPass.edit_ID + "'";
                        SqlCommand comm2 = new SqlCommand(sql2, conexion2);
                        SqlDataReader nwReader2 = comm2.ExecuteReader();
                    }
                    MessageBox.Show("Registro eliminado con exito.", "Editar", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al editar información.\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }
            }
        }
    }
}
