using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Web;
using System.Windows.Forms;

namespace GetName
{
    public partial class CrearPRF : Form
    {
        public CrearPRF()
        {
            InitializeComponent();
        }

        private static StringBuilder output = new StringBuilder();

        private void CMD(String command, Boolean msj=true)
        {
            output.Clear();
            ProcessStartInfo startInfo = new ProcessStartInfo("cmd.exe")
            {
                WindowStyle = ProcessWindowStyle.Hidden,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };

            startInfo.Arguments = @"/C "+command;
            Process process = Process.Start(startInfo);
            process.OutputDataReceived += new DataReceivedEventHandler((_sender, _e) =>
            {
                if (!String.IsNullOrEmpty(_e.Data))
                {
                    this.Invoke(new Action(() =>
                    {
                        richTextBox1.Text += _e.Data + "\n\n";
                    }));
                }
            });
            process.BeginOutputReadLine();
            process.WaitForExit();
        }

        private void CrearPRF_Load(object sender, EventArgs e)
        {
            //this.TopMost = true;
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            this.Invoke(new Action(() => { richTextBox1.Text = "Creando perfil de Outlook (" + VerPass.edit_nombre + ")..."; }));
 
            string path = Path.GetTempPath() + VerPass.edit_ID + ".PRF";
            //string path = @"c:\temp\UNNE.PRF";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            using (StreamWriter sw = File.CreateText(path))
            {
                this.Invoke(new Action(() => { richTextBox1.Text += "\n\nEscribiendo parametros de configuración..."; }));
                sw.WriteLine(";Automatically generated PRF file from the Microsoft Office Customization and Installation Wizard\r\r; **************************************************************\r; Section 1 - Profile Defaults\r; **************************************************************\r\r[General]\rCustom=1\rProfileName=UNNE\rDefaultProfile=Yes\rOverwriteProfile=Yes\rModifyDefaultProfileIfPresent=false\rDefaultStore=Service0\r\r; **************************************************************\r; Section 2 - Services in Profile\r; **************************************************************\r\r[Service List]\r;ServiceX=Microsoft Outlook Client\r\r;***************************************************************\r; Section 3 - List of internet accounts\r;***************************************************************\r\r[Internet Account List]\rAccount1=I_Mail\r\r;***************************************************************\r; Section 4 - Default values for each service.\r;***************************************************************\r\r;[ServiceX]\r;FormDirectoryPage=\r;-- The URL of Exchange Web Services Form Directory page used to create Web forms.\r;WebServicesLocation=\r;-- The URL of Exchange Web Services page used to display unknown forms.\r;ComposeWithWebServices=\r;-- Set to true to use Exchange Web Services to compose forms.\r;PromptWhenUsingWebServices=\r;-- Set to true to use Exchange Web Services to display unknown forms.\r;OpenWithWebServices=\r;-- Set to true to prompt user before opening unknown forms when using Exchange Web Services.\r\r;***************************************************************\r; Section 5 - Values for each internet account.\r;***************************************************************\r");
                sw.WriteLine("[Account1]");
                sw.WriteLine("UniqueService=No");
                sw.WriteLine("AccountName=" + VerPass.edit_nombre);
                sw.WriteLine("POP3Server=mail2.unne.com.mx");
                sw.WriteLine("SMTPServer=mail2.unne.com.mx");
                sw.WriteLine("POP3UserName=unne\\" + VerPass.edit_usuario);
                sw.WriteLine("EmailAddress=" + VerPass.edit_correo);
                sw.WriteLine("POP3UseSPA=0");
                sw.WriteLine("DisplayName=" + VerPass.edit_nombre);
                sw.WriteLine("ReplyEMailAddress=");
                sw.WriteLine("SMTPUseAuth=1");
                sw.WriteLine("SMTPAuthMethod=0");
                sw.WriteLine("ConnectionType=0");
                sw.WriteLine("LeaveOnServer=0x0");
                sw.WriteLine("POP3UseSSL=1");
                sw.WriteLine("ConnectionOID=MyConnection");
                sw.WriteLine("POP3Port=995");
                sw.WriteLine("ServerTimeOut=60");
                sw.WriteLine("SMTPPort=587");
                sw.WriteLine("SMTPSecureConnection=0");
                sw.WriteLine("DefaultAccount=TRUE");
                sw.WriteLine("");
                sw.WriteLine(";***************************************************************\r; Section 6 - Mapping for profile properties\r;***************************************************************\r\r[Microsoft Exchange Server]\rServiceName=MSEMS\rMDBGUID=5494A1C0297F101BA58708002B2A2517\rMailboxName=PT_STRING8,0x6607\rHomeServer=PT_STRING8,0x6608\rOfflineAddressBookPath=PT_STRING8,0x660E\rOfflineFolderPathAndFilename=PT_STRING8,0x6610\r\r[Exchange Global Section]\rSectionGUID=13dbb0c8aa05101a9bb000aa002fc45a\rMailboxName=PT_STRING8,0x6607\rHomeServer=PT_STRING8,0x6608\rConfigFlags=PT_LONG,0x6601\rRPCoverHTTPflags=PT_LONG,0x6623\rRPCProxyServer=PT_UNICODE,0x6622\rRPCProxyPrincipalName=PT_UNICODE,0x6625\rRPCProxyAuthScheme=PT_LONG,0x6627\rAccountName=PT_UNICODE,0x6620\r\r[Microsoft Mail]\rServiceName=MSFS\rServerPath=PT_STRING8,0x6600\rMailbox=PT_STRING8,0x6601\rPassword=PT_STRING8,0x67f0\rRememberPassword=PT_BOOLEAN,0x6606\rConnectionType=PT_LONG,0x6603\rUseSessionLog=PT_BOOLEAN,0x6604\rSessionLogPath=PT_STRING8,0x6605\rEnableUpload=PT_BOOLEAN,0x6620\rEnableDownload=PT_BOOLEAN,0x6621\rUploadMask=PT_LONG,0x6622\rNetBiosNotification=PT_BOOLEAN,0x6623\rNewMailPollInterval=PT_STRING8,0x6624\rDisplayGalOnly=PT_BOOLEAN,0x6625\rUseHeadersOnLAN=PT_BOOLEAN,0x6630\rUseLocalAdressBookOnLAN=PT_BOOLEAN,0x6631\rUseExternalToHelpDeliverOnLAN=PT_BOOLEAN,0x6632\rUseHeadersOnRAS=PT_BOOLEAN,0x6640\rUseLocalAdressBookOnRAS=PT_BOOLEAN,0x6641\rUseExternalToHelpDeliverOnRAS=PT_BOOLEAN,0x6639\rConnectOnStartup=PT_BOOLEAN,0x6642\rDisconnectAfterRetrieveHeaders=PT_BOOLEAN,0x6643\rDisconnectAfterRetrieveMail=PT_BOOLEAN,0x6644\rDisconnectOnExit=PT_BOOLEAN,0x6645\rDefaultDialupConnectionName=PT_STRING8,0x6646\rDialupRetryCount=PT_STRING8,0x6648\rDialupRetryDelay=PT_STRING8,0x6649\r\r[Personal Folders]\rServiceName=MSPST MS\rName=PT_STRING8,0x3001\rPathAndFilenameToPersonalFolders=PT_STRING8,0x6700 \rRememberPassword=PT_BOOLEAN,0x6701\rEncryptionType=PT_LONG,0x6702\rPassword=PT_STRING8,0x6703\r\r[Unicode Personal Folders]\rServiceName=MSUPST MS\rName=PT_UNICODE,0x3001\rPathAndFilenameToPersonalFolders=PT_STRING8,0x6700 \rRememberPassword=PT_BOOLEAN,0x6701\rEncryptionType=PT_LONG,0x6702\rPassword=PT_STRING8,0x6703\r\r[Outlook Address Book]\rServiceName=CONTAB\r\r[LDAP Directory]\rServiceName=EMABLT\rServerName=PT_STRING8,0x6600\rUserName=PT_STRING8,0x6602\rUseSSL=PT_BOOLEAN,0x6613\rUseSPA=PT_BOOLEAN,0x6615\rEnableBrowsing=PT_BOOLEAN,0x6622\rDisplayName=PT_STRING8,0x3001\rConnectionPort=PT_STRING8,0x6601\rSearchTimeout=PT_STRING8,0x6607\rMaxEntriesReturned=PT_STRING8,0x6608\rSearchBase=PT_STRING8,0x6603\rCheckNames=PT_STRING8,0x6624\rDefaultSearch=PT_LONG,0x6623\r\r[Microsoft Outlook Client]\rSectionGUID=0a0d020000000000c000000000000046\rFormDirectoryPage=PT_STRING8,0x0270\rWebServicesLocation=PT_STRING8,0x0271\rComposeWithWebServices=PT_BOOLEAN,0x0272\rPromptWhenUsingWebServices=PT_BOOLEAN,0x0273\rOpenWithWebServices=PT_BOOLEAN,0x0274\rCachedExchangeMode=PT_LONG,0x041f\rCachedExchangeSlowDetect=PT_BOOLEAN,0x0420\r\r[Personal Address Book]\rServiceName=MSPST AB\rNameOfPAB=PT_STRING8,0x001e3001\rPathAndFilename=PT_STRING8,0x001e6600\rShowNamesBy=PT_LONG,0x00036601\r\r; ************************************************************************\r; Section 7 - Mapping for internet account properties.  DO NOT MODIFY.\r; ************************************************************************\r\r[I_Mail]\rAccountType=POP3\r;--- POP3 Account Settings ---\rAccountName=PT_UNICODE,0x0002\rDisplayName=PT_UNICODE,0x000B\rEmailAddress=PT_UNICODE,0x000C\r;--- POP3 Account Settings ---\rPOP3Server=PT_UNICODE,0x0100\rPOP3UserName=PT_UNICODE,0x0101\rPOP3UseSPA=PT_LONG,0x0108\rOrganization=PT_UNICODE,0x0107\rReplyEmailAddress=PT_UNICODE,0x0103\rPOP3Port=PT_LONG,0x0104\rPOP3UseSSL=PT_LONG,0x0105\r; --- SMTP Account Settings ---\rSMTPServer=PT_UNICODE,0x0200\rSMTPUseAuth=PT_LONG,0x0203\rSMTPAuthMethod=PT_LONG,0x0208\rSMTPUserName=PT_UNICODE,0x0204\rSMTPUseSPA=PT_LONG,0x0207\rConnectionType=PT_LONG,0x000F\rConnectionOID=PT_UNICODE,0x0010\rSMTPPort=PT_LONG,0x0201\rSMTPSecureConnection=PT_LONG,0x020A\rServerTimeOut=PT_LONG,0x0209\rLeaveOnServer=PT_LONG,0x1000\r\r[IMAP_I_Mail]\rAccountType=IMAP\r;--- IMAP Account Settings ---\rAccountName=PT_UNICODE,0x0002\rDisplayName=PT_UNICODE,0x000B\rEmailAddress=PT_UNICODE,0x000C\r;--- IMAP Account Settings ---\rIMAPServer=PT_UNICODE,0x0100\rIMAPUserName=PT_UNICODE,0x0101\rIMAPUseSPA=PT_LONG,0x0108\rOrganization=PT_UNICODE,0x0107\rReplyEmailAddress=PT_UNICODE,0x0103\rIMAPPort=PT_LONG,0x0104\rIMAPUseSSL=PT_LONG,0x0105\r; --- SMTP Account Settings ---\rSMTPServer=PT_UNICODE,0x0200\rSMTPUseAuth=PT_LONG,0x0203\rSMTPAuthMethod=PT_LONG,0x0208\rSMTPUserName=PT_UNICODE,0x0204\rSMTPUseSPA=PT_LONG,0x0207\rConnectionType=PT_LONG,0x000F\rConnectionOID=PT_UNICODE,0x0010\rSMTPPort=PT_LONG,0x0201\rSMTPSecureConnection=PT_LONG,0x020A\rServerTimeOut=PT_LONG,0x0209\rCheckNewImap=PT_LONG,0x1100\rRootFolder=PT_UNICODE,0x1101\rAccount=PT_UNICODE,0x0002\rHttpServer=PT_UNICODE,0x0100\rUserName=PT_UNICODE,0x0101\rOrganization=PT_UNICODE,0x0107\rUseSPA=PT_LONG,0x0108\rTimeOut=PT_LONG,0x0209\rReply=PT_UNICODE,0x0103\rEmailAddress=PT_UNICODE,0x000C\rFullName=PT_UNICODE,0x000B\rConnection Type=PT_LONG,0x000F\rConnectOID=PT_UNICODE,0x0010");
            }
            this.Invoke(new Action(() => { richTextBox1.Text += "\n\nEliminando perfil creado...\n\n"; }));


            CMD("taskkill /im Outlook.exe");
            CMD("timeout /t 3 /nobreak");
            CMD("echo Importando perfil (OUTLOOK SE CERRARA AUTOMATICAMENTE)...");
            CMD("echo ¡¡NO CERRAR ESTA VENTANA!!");
            CMD("Start OUTLOOK.EXE /importprf \""+path+"\"");
            CMD("timeout /t 5 /nobreak");
            CMD("taskkill /im Outlook.exe");
            CMD("timeout /t 2 /nobreak");
            CMD("echo Configurando registro...");
            CMD("reg add \"HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\Outlook\\Profiles\\UNNE\\9375CFF0413111d3B88A00104B2A6676\\00000001\" /v \"POP3 Remember Passwd\" /t REG_DWORD /d 0 /f ");
            CMD("timeout /t 2 /nobreak");
            CMD("Start OUTLOOK.EXE");
            CMD("timeout /t 2 /nobreak");
            CMD("echo EN OUTLOOK PRESIONE ENVIAR Y RECIBIR Y PEGUE LA CONTRASEÑA Y PUEDE CERRAR ESTA VENTANA!");
            File.Delete(path);
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
        }
    }
}
