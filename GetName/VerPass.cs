using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace GetName
{
    public partial class VerPass : Form
    {
        /*
V1.7.7
    - Se agrega opcion de revisar si la cuenta sigue activa con clic derecho.
V1.7.6
    - Se agrega opcion de configurar correo desde el mismo programa.
V1.7.5
- Se encripta las contraseñas en la base de datos

V1.7.4
- Se deshabilita la opcion de copiar celdas.
- Se puede copiar toda la fila haciendo clic sobre cualquier celda y precionando la tecla C.

V1.73
- Se agrega boton de actualizar
        */
        private String versiontext = "1.7.7";
        private String version = "070b7fa645faaa4e7a8e0e90af0ce82611c29fe7";
        public static String conexionsqllast = "server=148.223.153.37,5314; database=InfEq;User ID=eordazs;Password=Corpame*2013; integrated security = false ; MultipleActiveResultSets=True";
        public static String buscar;

        public static Boolean copiar = false;

        public VerPass()
        {
            InitializeComponent();
        }
        public static List<String[]> lista = new List<String[]>();
        
        private void VerPass_Load(object sender, EventArgs e)
        {
            dataGridView1.MultiSelect = copiar;
            if(backgroundWorker1.IsBusy != true)
            {
                label1.Text = "";
                lista.Clear();
                dataGridView1.Rows.Clear();
                pictureBox1.Enabled = true;
                pictureBox1.BringToFront();
                pictureBox1.Visible = true;
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void TextBox1_KeyUp(object sender, KeyEventArgs e)
        {
            dataGridView1.Rows.Clear();
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                foreach (String[] empleado in lista)
                {
                    String nmayus = empleado[0].ToString().ToUpper();
                    if (nmayus.Contains(textBox1.Text.ToUpper()))
                    {
                        dataGridView1.Rows.Add(empleado[0], empleado[1], empleado[2], empleado[3], empleado[4], empleado[5]);
                    }
                }
            }
            else
            {
                foreach (String[] empleado in lista)
                {
                    dataGridView1.Rows.Add(empleado[0], empleado[1], empleado[2], empleado[3], empleado[4], empleado[5]);
                }
            }
            buscar = textBox1.Text;
        }

        private void portapapeles(String pegar, String msj="")
        {
            label1.Text = "Portapapeles: " + pegar;
            if (!String.IsNullOrEmpty(msj))
            {
                label1.Text = msj;
            }
            Clipboard.SetText(pegar);
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == this.dataGridView1.Columns["nombre"].Index && e.RowIndex != -1)
            {
                portapapeles(dataGridView1.Rows[e.RowIndex].Cells["nombre"].Value.ToString());
            }
            if (e.ColumnIndex == this.dataGridView1.Columns["usuario"].Index && e.RowIndex != -1)
            {
                portapapeles("unne\\"+dataGridView1.Rows[e.RowIndex].Cells["usuario"].Value.ToString());
            }
            if (e.ColumnIndex == this.dataGridView1.Columns["correo"].Index && e.RowIndex != -1)
            {
                portapapeles(dataGridView1.Rows[e.RowIndex].Cells["correo"].Value.ToString());
            }
            if (e.ColumnIndex == this.dataGridView1.Columns["password"].Index && e.RowIndex != -1)
            {
                portapapeles(dataGridView1.Rows[e.RowIndex].Cells["password"].Value.ToString());
            }
        }

        private void CopiarDNSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            portapapeles("mail2.unne.com.mx");
        }

        private void CerrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void GuardarCorreoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.ShowDialog();
            VerPass_Load(sender, e);
        }

        private static bool Verificar_Correo(string username, string password)
        {
            var service = new ExchangeService(ExchangeVersion.Exchange2013);
            service.Credentials = new WebCredentials(username, password);
            service.Url = new Uri("https://mail2.unne.com.mx/EWS/Exchange.asmx");

            Console.WriteLine("Verificando: " + username);

            try
            {
                service.GetAppManifests();
                return true;
            }
            catch
            {
                return false;
            }

            return false;
        }

        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                using (SqlConnection conexion2 = new SqlConnection(conexionsqllast))
                {
                    conexion2.Open();
                    String sql2 = "SELECT (SELECT valor FROM Configuracion WHERE nombre='GetName_Version') as version,valor FROM Configuracion WHERE nombre='GetName_hash'";
                    SqlCommand comm2 = new SqlCommand(sql2, conexion2);
                    SqlDataReader nwReader2 = comm2.ExecuteReader();
                    if (nwReader2.Read())
                    {
                        if (nwReader2["valor"].ToString() != version)
                        {
                            MessageBox.Show("No se puede inciar sesion porque hay una nueva version.\n\nNueva Version: "+ nwReader2["version"].ToString() + "\nVersion actual: "+versiontext+"\n\nEl programa se cerrara.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            System.Windows.Forms.Application.Exit();
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se puedo verificar la version del programa.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        System.Windows.Forms.Application.Exit();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error en validar la version del programa\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
                return;
            }
            
            try
            {
                using (SqlConnection conexion = new SqlConnection(conexionsqllast))
                {
                    conexion.Open();
                    SqlCommand comm = new SqlCommand("SELECT * FROM [InfEq].[dbo].[GetName] ORDER BY id DESC", conexion);
                    SqlDataReader nwReader = comm.ExecuteReader();
                    while (nwReader.Read())
                    {
                        String[] n = new String[7];
                        n[0] = nwReader["nombre"].ToString();
                        n[1] = nwReader["usuario"].ToString();
                        n[2] = nwReader["correo"].ToString();
                        n[3] = Desencriptar(nwReader["password"].ToString());
                        n[4] = nwReader["fecha"].ToString();
                        n[5] = nwReader["id"].ToString();

                        /*
                        n[6] = "0";
                        if (Verificar_Correo("unne\\"+n[1],n[3]) == true)
                        {
                            n[6] = "1";
                        }
                        */

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

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if(e.Cancelled)
            {
                System.Windows.Forms.Application.Exit();
            }
            if(lista.Any())
            {
                int contarceldas = 0;
                foreach(String[] n in lista)
                {
                    if (!String.IsNullOrEmpty(buscar))
                    {
                        textBox1.Text = buscar;
                        String nmayus = n[0].ToString().ToUpper();
                        if (nmayus.Contains(buscar.ToUpper()))
                        {
                            dataGridView1.Rows.Add(n[0], n[1], n[2], n[3], n[4], n[5]);
                        }
                    }
                    else
                    {
                        dataGridView1.Rows.Add(n[0], n[1], n[2], n[3], n[4], n[5]);
                    }
                    /*
                    if(n[6] == "0")
                    {
                        dataGridView1.Rows[contarceldas].DefaultCellStyle.BackColor = Color.LightSalmon;
                    }
                    else
                    {
                        dataGridView1.Rows[contarceldas].DefaultCellStyle.BackColor = Color.PaleGreen;
                    }
                    */
                    contarceldas++;
                }
            }
            pictureBox1.Visible = false;
            pictureBox1.Enabled = false;
        }

        private void reenviarCorreoToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            String Nombre = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            String correo = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            String usuario = "unne\\"+dataGridView1.CurrentRow.Cells[1].Value.ToString();
            String pass = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            DialogResult reenviar = MessageBox.Show("Reenviar contraseña de\n" + Nombre + "?", "Ver Contraseñas", MessageBoxButtons.YesNo,MessageBoxIcon.Question);
            if(reenviar == DialogResult.Yes)
            {
                List<string> lstAllRecipients = new List<string>();
                //Below is hardcoded - can be replaced with db data
                lstAllRecipients.Add(correo);

                Outlook.Application outlookApp = new Outlook.Application();
                Outlook._MailItem oMailItem = (Outlook._MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Inspector oInspector = oMailItem.GetInspector;

                // Recipient
                Outlook.Recipients oRecips = (Outlook.Recipients)oMailItem.Recipients;
                foreach (String recipient in lstAllRecipients)
                {
                    Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(recipient);
                    oRecip.Resolve();
                }

                //Copia de correo
                /*
                Outlook.Recipient oCCRecip = oRecips.Add(copiacorreo);

                oCCRecip.Type = (int)Outlook.OlMailRecipientType.olBCC;
                oCCRecip.Resolve();
                */

                        oMailItem.Subject = "Contraseña de correo";

                String FirmaBody = oMailItem.HTMLBody;

                //Body para Edson
                oMailItem.HTMLBody = "<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><meta http-equiv=Content-Type content=\"text/html; charset=iso-8859-1\"><meta name=Generator content=\"Microsoft Word 15 (filtered medium)\"><style><!--/* Font Definitions */@font-face	{font-family:\"Cambria Math\";	panose-1:2 4 5 3 5 4 6 3 2 4;}@font-face	{font-family:Calibri;	panose-1:2 15 5 2 2 2 4 3 2 4;}/* Style Definitions */p.MsoNormal, li.MsoNormal, div.MsoNormal	{margin:0cm;	margin-bottom:.0001pt;	font-size:11.0pt;	font-family:\"Calibri\",sans-serif;	mso-fareast-language:EN-US;}a:link, span.MsoHyperlink	{mso-style-priority:99;	color:#0563C1;	text-decoration:underline;}a:visited, span.MsoHyperlinkFollowed	{mso-style-priority:99;	color:#954F72;	text-decoration:underline;}span.EstiloCorreo17	{mso-style-type:personal-compose;	font-family:\"Calibri\",sans-serif;	color:#2F5496;}.MsoChpDefault	{mso-style-type:export-only;	font-family:\"Calibri\",sans-serif;	mso-fareast-language:EN-US;}@page WordSection1	{size:612.0pt 792.0pt;	margin:70.85pt 3.0cm 70.85pt 3.0cm;}div.WordSection1	{page:WordSection1;}--></style><!--[if gte mso 9]><xml><o:shapedefaults v:ext=\"edit\" spidmax=\"1026\" /></xml><![endif]--><!--[if gte mso 9]><xml><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\" /></o:shapelayout></xml><![endif]--></head><body lang=ES-MX link=\"#0563C1\" vlink=\"#954F72\"><div class=WordSection1><p class=MsoNormal><span style='color:#2F5496'>Buen día.<o:p></o:p></span></p><p class=MsoNormal><span style='color:#2F5496'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span style='color:#2F5496'>Por medio de la presente te hacemos llegar tu usuario y contraseña para el uso del correo electrónico corporativo. Es muy importante que utilices correctamente esta información considerando las siguientes políticas:<o:p></o:p></span></p><p class=MsoNormal><span style='color:#2F5496'>&#8226;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; El uso de internet es una herramienta de trabajo para el uso de actividades de la empresa.<o:p></o:p></span></p><p class=MsoNormal><span style='color:#2F5496'>&#8226;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Está estrictamente prohibido escuchar música&nbsp; y ver videos en internet, así como cualquier otra actividad que consuma recursos de internet.<o:p></o:p></span></p><p class=MsoNormal><span style='color:#2F5496'>&#8226;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Se prohíbe la instalación de programas no autorizados por la empresa y en caso de requerir la instalación de un programa, favor de apoyarse con el área de sistemas con la finalidad de no instalar software basura que viene con los programas gratuitos.<o:p></o:p></span></p><p class=MsoNormal><span style='color:#2F5496'>&#8226;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; La utilización del usuario y contraseña son de uso PERSONAL, es decir, que tú eres el responsable de cualquier solicitud que se haga en el sistema con ese usuario, por lo tanto no debes prestarlo a nadie.<o:p></o:p></span></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse'><tr><td width=198 valign=top style='width:148.45pt;border:solid windowtext 1.0pt;background:#A8D08D;padding:0cm 0cm 0cm 0cm'><p class=MsoNormal align=center style='text-align:center'><b>CORREO ELECTRÓNICO<o:p></o:p></b></p></td><td width=106 style='width:79.65pt;border:solid windowtext 1.0pt;border-left:none;background:#A8D08D;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b>USUARIO<o:p></o:p></b></p></td><td width=151 style='width:4.0cm;border:solid windowtext 1.0pt;border-left:none;background:#A8D08D;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b>CONTRASEÑA<o:p></o:p></b></p></td></tr><tr><td width=198 valign=top style='width:148.45pt;border:solid windowtext 1.0pt;border-top:none;padding:0cm 0cm 0cm 0cm'><p class=MsoNormal align=center style='text-align:center'><a href=\"mailto:" + correo + "\">" + correo + "</a><o:p></o:p></p></td><td width=106 style='width:79.65pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'>" + usuario + "<o:p></o:p></p></td><td width=151 style='width:4.0cm;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'>" + pass + "<o:p></o:p></p></td></tr></table>";

                //Body para Daniel
                //oMailItem.HTMLBody = "<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><meta http-equiv=Content-Type content=\"text/html; charset=iso-8859-1\"><meta name=Generator content=\"Microsoft Word 15 (filtered medium)\"><style><!--/* Font Definitions */@font-face	{font-family:\"Cambria Math\";	panose-1:2 4 5 3 5 4 6 3 2 4;}@font-face	{font-family:Calibri;	panose-1:2 15 5 2 2 2 4 3 2 4;}/* Style Definitions */p.MsoNormal, li.MsoNormal, div.MsoNormal	{margin:0cm;	margin-bottom:.0001pt;	font-size:11.0pt;	font-family:\"Calibri\",sans-serif;	mso-fareast-language:EN-US;}a:link, span.MsoHyperlink	{mso-style-priority:99;	color:#0563C1;	text-decoration:underline;}a:visited, span.MsoHyperlinkFollowed	{mso-style-priority:99;	color:#954F72;	text-decoration:underline;}span.EstiloCorreo17	{mso-style-type:personal-compose;	font-family:\"Calibri\",sans-serif;	color:#2F5496;}.MsoChpDefault	{mso-style-type:export-only;	font-size:10.0pt;	font-family:\"Calibri\",sans-serif;	mso-fareast-language:EN-US;}@page WordSection1	{size:612.0pt 792.0pt;	margin:70.85pt 3.0cm 70.85pt 3.0cm;}div.WordSection1	{page:WordSection1;}--></style><!--[if gte mso 9]><xml><o:shapedefaults v:ext=\"edit\" spidmax=\"1026\" /></xml><![endif]--><!--[if gte mso 9]><xml><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\" /></o:shapelayout></xml><![endif]--></head><body lang=ES-MX link=\"#0563C1\" vlink=\"#954F72\"><div class=WordSection1><p class=MsoNormal>Buen día.<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>Por medio de la presente te hacemos llegar tu usuario y contraseña para el uso del correo electrónico corporativo. Es muy importante que utilices correctamente esta información considerando las siguientes políticas:<o:p></o:p></p><p class=MsoNormal>&#8226;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; El uso de internet es una herramienta de trabajo para el uso de actividades de la empresa.<o:p></o:p></p><p class=MsoNormal>&#8226;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Está estrictamente prohibido escuchar música&nbsp; y ver videos en internet, así como cualquier otra actividad que consuma recursos de internet.<o:p></o:p></p><p class=MsoNormal>&#8226;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Se prohíbe la instalación de programas no autorizados por la empresa y en caso de requerir la instalación de un programa, favor de apoyarse con el área de sistemas con la finalidad de no instalar software basura que viene con los programas gratuitos.<o:p></o:p></p><p class=MsoNormal>&#8226;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; La utilización del usuario y contraseña son de uso PERSONAL, es decir, que tú eres el responsable de cualquier solicitud que se haga en el sistema con ese usuario, por lo tanto no debes prestarlo a nadie.<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse'><tr><td width=198 valign=top style='width:148.45pt;border:solid windowtext 1.0pt;background:#A8D08D;padding:0cm 0cm 0cm 0cm'><p class=MsoNormal align=center style='text-align:center'><b>CORREO ELECTRÓNICO<o:p></o:p></b></p></td><td width=106 style='width:79.65pt;border:solid windowtext 1.0pt;border-left:none;background:#A8D08D;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b>USUARIO<o:p></o:p></b></p></td><td width=151 style='width:4.0cm;border:solid windowtext 1.0pt;border-left:none;background:#A8D08D;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b>CONTRASEÑA<o:p></o:p></b></p></td></tr><tr><td width=198 valign=top style='width:148.45pt;border:solid windowtext 1.0pt;border-top:none;padding:0cm 0cm 0cm 0cm'><p class=MsoNormal align=center style='text-align:center'><a href=\"mailto:" + correo + "\">" + correo + "</a><o:p></o:p></p></td><td width=106 style='width:79.65pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'>" + usuario + "<o:p></o:p></p></td><td width=151 style='width:4.0cm;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'>" + pass.Text + "<o:p></o:p></p></td></tr></table><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>";

                oMailItem.HTMLBody += FirmaBody;
                oMailItem.Display(true);
            }
        }

        public static string Encriptar(String textToEncrypt)
        {
            try
            {
                //string textToEncrypt = "EdsonOrdaz";;
                string ToReturn = "";
                string publickey = "12345678";
                string secretkey = "87654321";
                byte[] secretkeyByte = { };
                secretkeyByte = System.Text.Encoding.UTF8.GetBytes(secretkey);
                byte[] publickeybyte = { };
                publickeybyte = System.Text.Encoding.UTF8.GetBytes(publickey);
                MemoryStream ms = null;
                CryptoStream cs = null;
                byte[] inputbyteArray = System.Text.Encoding.UTF8.GetBytes(textToEncrypt);
                using (DESCryptoServiceProvider des = new DESCryptoServiceProvider())
                {
                    ms = new MemoryStream();
                    cs = new CryptoStream(ms, des.CreateEncryptor(publickeybyte, secretkeyByte), CryptoStreamMode.Write);
                    cs.Write(inputbyteArray, 0, inputbyteArray.Length);
                    cs.FlushFinalBlock();
                    ToReturn = Convert.ToBase64String(ms.ToArray());
                }
                return ToReturn;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex.InnerException);
            }
        }

        public static string Desencriptar(String textToDecrypt)
        {
            try
            {
                //string textToDecrypt = "6aQ+vnbOIuOZsJPfOgWoVA==";
                string ToReturn = "";
                string publickey = "12345678";
                string secretkey = "87654321";
                byte[] privatekeyByte = { };
                privatekeyByte = System.Text.Encoding.UTF8.GetBytes(secretkey);
                byte[] publickeybyte = { };
                publickeybyte = System.Text.Encoding.UTF8.GetBytes(publickey);
                MemoryStream ms = null;
                CryptoStream cs = null;
                byte[] inputbyteArray = new byte[textToDecrypt.Replace(" ", "+").Length];
                inputbyteArray = Convert.FromBase64String(textToDecrypt.Replace(" ", "+"));
                using (DESCryptoServiceProvider des = new DESCryptoServiceProvider())
                {
                    ms = new MemoryStream();
                    cs = new CryptoStream(ms, des.CreateDecryptor(publickeybyte, privatekeyByte), CryptoStreamMode.Write);
                    cs.Write(inputbyteArray, 0, inputbyteArray.Length);
                    cs.FlushFinalBlock();
                    Encoding encoding = Encoding.UTF8;
                    ToReturn = encoding.GetString(ms.ToArray());
                }
                return ToReturn;
            }
            catch (Exception ae)
            {
                throw new Exception(ae.Message, ae.InnerException);
            }
        }



        public static String edit_ID;
        public static String edit_nombre;
        public static String edit_usuario;
        public static String edit_correo;
        public static String edit_pass;
        
        private void editarDatosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            edit_ID = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            edit_nombre = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            edit_correo = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            edit_usuario = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            edit_pass = dataGridView1.CurrentRow.Cells[3].Value.ToString();

            Editar edit = new Editar();
            edit.ShowDialog();
            VerPass_Load(sender, e);
        }

        private void actualizarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VerPass_Load(sender, e);
        }

        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.KeyCode==Keys.C)
            {
                String Nombre = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                String correo = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                String usuario = "unne\\" + dataGridView1.CurrentRow.Cells[1].Value.ToString();
                String pass = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                String fecha = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                portapapeles("Nombre: " + Nombre + "\nUsuario: " + usuario + "\nContraseña: " + pass + "\nCorreo: " + correo+"\nÚltima modificación: "+fecha, "Datos del usuario copiados.");
            }
        }


        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            edit_ID = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            edit_nombre = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            edit_correo = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            edit_usuario = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            edit_pass = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            DialogResult result = MessageBox.Show("¿Configurar en Outlook el correo de "+edit_nombre+"?", "Configurar Correo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                portapapeles(dataGridView1.Rows[e.RowIndex].Cells["password"].Value.ToString());
                CrearPRF crearPRF = new CrearPRF();
                crearPRF.ShowDialog();
            }
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                label1.Text = "Verificando...";
                String u = "unne\\"+dataGridView1.Rows[e.RowIndex].Cells["usuario"].Value.ToString();
                String p = dataGridView1.Rows[e.RowIndex].Cells["password"].Value.ToString();
                if (Verificar_Correo(u, p))
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.PaleGreen;
                    label1.Text = "Usuario y contraseña correctos";
                }
                else
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightSalmon;
                    label1.Text = "Contraseña incorrecta";
                }
            }
        }
    }
}
