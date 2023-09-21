using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Net;
using System.Threading;
using System.Collections;
using System.Net.Security;
using System.Net;
using System.Net.Mail;

namespace CachaYEnvia
{
    /* CONTROL DE CAMBIOS
     *20131106  Este programa al momento que se ejecute, disparado por una tarea programada, buscará en un conjunto de carpetas (definidas en el archivo config) 
     *          un archivo que entre su nombre (la mascara del config), tiene la fecha del dia anterior. 
     *          cada uno de los archivos coincidentes será anexado a un correo y enviado a los diferentes destinatarios via correo electrónico.
     *          Al terminar de enviar el correo el programa se cerrará.
     *          
     *          El progrmaa avisa a destinatarios de sistemas por si algún reporte no fue creado.
     *          El programa envia a cada una de las marcas (correos por marca), los reportes que le pertenecen a esa marca ese día.
     * 
     *          20131216

                Es necesario que se ejecute en dos servidores distintos (A y B), y que solo envié un correo electrónico conteniendo los archivos de todas las agencias.
                Por tanto el proceso que se ejecuta al final en el tiempo (en el servidor B), se encargará de recolectar los archivos del servidor A, y  enviar el archivo final.
     * 
     *          20140322
     *          Se anexa desde que ip se envia el mensaje, se parametriza el EnableSsl del correo, se borra cualquier zip de la carpeta desde donde se ejecuta el aplicativo
     *         
     *          20140328 
     *          En ocasiones el archivo a enviar se genera en el mismo día de la ejecución, entonces si no encuentra el archivo de la fecha anterior, enviará el archivo de la fecha corriente, el más reciente.
     *          al otro día encontrará dos archivos de la fecha del día anterior --> debe enviar el del día anterior el más reciente.
     * */

    public partial class Envia : Form
    {
            Dictionary<string, string> dic_emailsxmarca = new Dictionary<string, string>();
            ArrayList arrTodasMarcas = new ArrayList();    

            string DirLocalArchivosZipear = ""; //Directorio usado para zipear los archivos a enviar via email
            string DirectorioObservar = ""; //Directorio donde se buscará la subcarpeta dinámica.            
            string IPRemoto = System.Configuration.ConfigurationSettings.AppSettings["IPRemoto"];
            string Usr=System.Configuration.ConfigurationSettings.AppSettings["Usr"];
            string Pass=System.Configuration.ConfigurationSettings.AppSettings["Pass"];
            string CarpetaRemota = System.Configuration.ConfigurationSettings.AppSettings["CarpetaRemota"];
            string TODASLASMarcasCSV = System.Configuration.ConfigurationSettings.AppSettings["TODASLASMarcasCSV"];  
            string EnviaCorreoPrincipal = System.Configuration.ConfigurationSettings.AppSettings["EnviaCorreoPrincipal"]; //indica si este executable enviará el correo principal de aviso de los archivos.

            string PreSubject = System.Configuration.ConfigurationSettings.AppSettings["PreSubject"]; //Se agrega al subject para efectos de pruebas.

            string Mascara=""; //El patrón que se buscará en el nombre de archivos a enviar.
            
            string DirectorioEnviados = "";
            string MinutosEjecucion = "0"; //Son los minutos que estará activo el Programa, 0 para que permanezca en ejecucion indefinidamente
            
            string MetodoDeEnvio = "CARPETACOMPARTIDA";
            
            string DirSincronizacionRemota = ""; //Donde dejará la GUIA para que Sincronizacion la tome.    
            //string NumeroSucursalEnvia = "";
            
            string MinutosEsperaLlenado = ""; //el tiempo que se deberá esperar para que llene el archivo zip             
            string ConnectionString = "";
            string MinutosEjecucionDespuesUltimoEnviado = ""; //Minutos que esperará para cerrar el aplicativo despues de la recepcion y envio del último zip.            

            string smtpserverhost = "";
            string smtpport="";
            string usrcredential="";
            string usrpassword="";
            string EnableSsl = "";                

            string emailsavisar = "";
            string IPLocal = "";        
            
            string LInferiorInicioTyCHoraNoc="";
            string LSuperiorInicioTyCHoraNoc = "";
            string HoraCerrarAplicacion = System.Configuration.ConfigurationSettings.AppSettings["HoraCerrarAplicacion"];

            string CarpetasCSV = System.Configuration.ConfigurationSettings.AppSettings["CarpetasCSV"]; //Las carpetas donde se buscarán los archivos.
            string MarcasCSV = System.Configuration.ConfigurationSettings.AppSettings["MarcasCSV"]; //Se corresponden en orden con las carpetas, son los nombres de las marcas de autos
            string emailsavisarNoHayNada = System.Configuration.ConfigurationSettings.AppSettings["emailsavisarNoHayNada"]; //avisa a estos destinatarios aquellas agencias que no se les encontró archivo el día de hoy.
            string emailsXMarca = System.Configuration.ConfigurationSettings.AppSettings["emailsXMarca"]; //a cada marca se le enviarán sus correspondientes archivos.


            DateTime hora_inicio = DateTime.Now;

            string[] Argumentos;

            #region Impersonacion en el servidor remoto
            [DllImport("advapi32.dll", SetLastError = true)]
            private static extern bool LogonUser(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
            private unsafe static extern int FormatMessage(int dwFlags, ref IntPtr lpSource, int dwMessageId, int dwLanguageId, ref String lpBuffer, int nSize, IntPtr* arguments);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern bool CloseHandle(IntPtr handle);

            [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public extern static bool DuplicateToken(IntPtr existingTokenHandle, int SECURITY_IMPERSONATION_LEVEL, ref IntPtr duplicateTokenHandle);

            // logon types
            const int LOGON32_LOGON_INTERACTIVE = 2;
            const int LOGON32_LOGON_NETWORK = 3;
            const int LOGON32_LOGON_NEW_CREDENTIALS = 9;

            // logon providers
            const int LOGON32_PROVIDER_DEFAULT = 0; //0
            const int LOGON32_PROVIDER_WINNT50 = 3; //3
            const int LOGON32_PROVIDER_WINNT40 = 2;
            const int LOGON32_PROVIDER_WINNT35 = 1;

            #region manejo de errores
            // GetErrorMessage formats and returns an error message
            // corresponding to the input errorCode.
            public unsafe static string GetErrorMessage(int errorCode)
            {
                int FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100;
                int FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200;
                int FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;

                int messageSize = 255;
                string lpMsgBuf = "";
                int dwFlags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;

                IntPtr ptrlpSource = IntPtr.Zero;
                IntPtr ptrArguments = IntPtr.Zero;

                int retVal = FormatMessage(dwFlags, ref ptrlpSource, errorCode, 0, ref lpMsgBuf, messageSize, &ptrArguments);
                if (retVal == 0)
                {
                    throw new ApplicationException(string.Format("Failed to format message for error code '{0}'.", errorCode));
                }

                return lpMsgBuf;
            }

            private static void RaiseLastError()
            {
                int errorCode = Marshal.GetLastWin32Error();
                string errorMessage = GetErrorMessage(errorCode);

                throw new ApplicationException(errorMessage);
            }

            #endregion


            #endregion

        public Envia(string[] args)
        {
            if (args.Length > 0)
                this.Argumentos = args;
            else
                this.Argumentos = "0".Split(',');

            InitializeComponent();
        }

        private void Envia_Load(object sender, EventArgs e)
        {
           this.DirectorioObservar = System.Configuration.ConfigurationSettings.AppSettings["DirectorioObservar"];
           this.IPRemoto = System.Configuration.ConfigurationSettings.AppSettings["IPRemoto"];
           this.Usr = System.Configuration.ConfigurationSettings.AppSettings["Usr"];
           this.Pass = System.Configuration.ConfigurationSettings.AppSettings["Pass"];
           //this.DirRemoto = System.Configuration.ConfigurationSettings.AppSettings["DirRemoto"];
           this.Mascara = System.Configuration.ConfigurationSettings.AppSettings["Mascara"];
           //this.EnviarArchivosDelaFecha = System.Configuration.ConfigurationSettings.AppSettings["EnviarArchivosDelaFecha"];
           this.DirectorioEnviados  = System.Configuration.ConfigurationSettings.AppSettings["DirectorioEnviados"];
           this.MetodoDeEnvio = System.Configuration.ConfigurationSettings.AppSettings["MetodoDeEnvio"];
           this.DirSincronizacionRemota = System.Configuration.ConfigurationSettings.AppSettings["DirSincronizacionRemota"];
           //this.NumeroSucursalEnvia = System.Configuration.ConfigurationSettings.AppSettings["NumeroSucursalEnvia"];
           this.MinutosEsperaLlenado = System.Configuration.ConfigurationSettings.AppSettings["MinutosEsperaLlenado"];
           this.DirLocalArchivosZipear = System.Configuration.ConfigurationSettings.AppSettings["DirLocalArchivosZipear"];
           this.ConnectionString = System.Configuration.ConfigurationSettings.AppSettings["ConnectionString"];
           this.MinutosEjecucionDespuesUltimoEnviado = System.Configuration.ConfigurationSettings.AppSettings["MinutosEjecucionDespuesUltimoEnviado"];

           this.smtpserverhost = System.Configuration.ConfigurationSettings.AppSettings["smtpserverhost"];
           this.smtpport = System.Configuration.ConfigurationSettings.AppSettings["smtpport"];
           this.usrcredential = System.Configuration.ConfigurationSettings.AppSettings["usrcredential"];
           this.usrpassword = System.Configuration.ConfigurationSettings.AppSettings["usrpassword"];
           this.EnableSsl = System.Configuration.ConfigurationSettings.AppSettings["EnableSsl"];
           this.PreSubject = System.Configuration.ConfigurationSettings.AppSettings["PreSubject"]; //Se agrega al subject para efectos de pruebas.

           this.emailsavisar = System.Configuration.ConfigurationSettings.AppSettings["emailsavisar"];
           this.IPLocal = System.Configuration.ConfigurationSettings.AppSettings["IPLocal"];

           this.LInferiorInicioTyCHoraNoc = System.Configuration.ConfigurationSettings.AppSettings["LInferiorInicioTyCHoraNoc"];
           this.LSuperiorInicioTyCHoraNoc = System.Configuration.ConfigurationSettings.AppSettings["LSuperiorInicioTyCHoraNoc"];


           //" value="NISSAN:luis.bonnet@grupoandrade.com.mx,bonnetalj@hotmail.com;GMC:luis.bonnet@grupoandrade.com.mx,bonnetalj@hotmail.com;FORD:luis.bonnet@grupoandrade.com.mx,bonnetalj@hotmail.com" />
            //llenamos el diccionario
           string[] arrAux = this.emailsXMarca.Split(';');
           foreach (string marca in arrAux)
           { 
             string solomarca = marca.Substring(0,marca.IndexOf(":"));
             string correos = marca.Substring(marca.IndexOf(":")+1);
             if (!this.dic_emailsxmarca.Keys.Contains(solomarca.Trim()))
             {
                 this.dic_emailsxmarca.Add(solomarca.Trim(), correos.Trim());
             }           
           }

           arrAux = this.TODASLASMarcasCSV.Split(',');
           foreach (string marca in arrAux)
           {
               if (!this.arrTodasMarcas.Contains(marca))
                   this.arrTodasMarcas.Add(marca.Trim());
           }

           //borramos cualquier zip que contenga la carpeta desde donde se genera el aplicativo
           string[] archivosborrar = Directory.GetFiles(Application.StartupPath, "*.zip");
           foreach (string archxborrar in archivosborrar)
           {
               File.Delete(archxborrar);
           }

           //Utilerias.LimpiaArchivoLog(Application.StartupPath + "\\Log.txt");
           Utilerias.WriteToLog(" ", " ", Application.StartupPath + "\\Log.txt");
           Utilerias.WriteToLog("Inicio de operaciones" , "Envia_Load", Application.StartupPath + "\\Log.txt");
                                          
               if (this.MinutosEjecucion.Trim() == "" )
               {
                   this.MinutosEjecucion = "0";
               }

               this.timer1.Enabled = true;
               this.timer1.Start();
                                 
                 int ArchivosEnviados = EnviaArchivos();                 
                  Utilerias.WriteToLog("Se han enviado " + ArchivosEnviados.ToString() + " Exitosamente", "Envia_Load", Application.StartupPath + "\\Log.txt");
                  Utilerias.WriteToLog("El aplicactivo se cierra al enviar archivos", "Envia_Load", Application.StartupPath + "\\Log.txt");
                  Application.Exit();                      
        }

                

        /// <summary>
        /// Recorre todas las marcas y busca en la carpeta a zippear si hay un archivo con ese nombre.
        /// </summary>
        /// <returns>El nombre de las agencias que no se encontró archivo separadas por coma</returns>
        public string ValidaLaYaExistencia()
        {
            string res = "";
            try
            {
                //por cada marca registrada en el arreglo de todas las marcas,
                //validamos si existe o no un archivo pdf en la carpeta de archivos por zippear
                foreach (string marca in arrTodasMarcas)
                {
                    string ArchivoBuscar = "*" + marca + "*.pdf";
                    string[] archivos = Directory.GetFiles(this.DirLocalArchivosZipear, ArchivoBuscar);

                    if (archivos.Length == 0)
                    {
                        res += "Para : " + marca + " No se generó el archivo: " + ArchivoBuscar.Trim() + ",";
                        Utilerias.WriteToLog("No se encontró el archivo: " + ArchivoBuscar + " en la carpeta a zippear: " + this.DirLocalArchivosZipear.Trim(), "ValidaLaYaExistencia", Application.StartupPath + "\\Log.txt");
                    }
                }                
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            return res;        
        }


        public string[] RegresaElUltimo(string[] archivos)
        {
            string[] res = new string[1];

            DateTime FechaAyer = DateTime.Today.AddDays(-1);
            DateTime fechahoramasalta = new DateTime(FechaAyer.Year,FechaAyer.Month,FechaAyer.Day,0,0,0); 

            foreach (string archivo in archivos)
            {
                FileInfo fi = new FileInfo(archivo);
                if (fi.CreationTime >= fechahoramasalta)
                {
                    res[0] = archivo.Trim();
                    fechahoramasalta = fi.CreationTime;
                }
            }
            return res;
        }

        /// <summary>
        /// Recorrerá todas las carpetas en busca de los archivos a Enviar.
        /// Cuando los encuentre los copia a una carpeta donde serán zipeados.
        /// 
        /// </summary>
        /// <returns>El numero de archivos zip transferidos.</returns>
        public int EnviaArchivos()
        {
            int res = 0;
            try
            {
                //borramos cualquier cosa que tenga la carpeta a comprimir
                string[] archivosborrar = Directory.GetFiles(this.DirLocalArchivosZipear, "*.*");
                foreach (string archxborrar in archivosborrar)
                {
                    File.Delete(archxborrar);                
                }

                //Construimos el nombre del archivo dinámico con la fecha actual
                DateTime FechaAyer = DateTime.Today.AddDays(-1);
                string strFechaAyer = FechaAyer.ToString("ddMMyyyy");
                string strFechaHoy = DateTime.Today.ToString("ddMMyyyy");

                string[] CarpetasObservar = this.CarpetasCSV.Split(',');
                string[] NombresMarcas = this.MarcasCSV.Split(',');

                int i = 0;
                int indicemarca = 0;
                //CierreDia_Conciliacion_Modulos vs contabilidad_31102013021026.pdf
                //this.Mascara = "CierreDia_Conciliacion_Modulos vs contabilidad_";
                string ArchivoBuscarAyer = this.Mascara.Trim() + strFechaAyer + "*.pdf";
                string ArchivoBuscarHoy = this.Mascara.Trim() + strFechaHoy + "*.pdf";


                foreach (string CarpetaObservar in CarpetasObservar)
                {
                    string[] archivos = Directory.GetFiles(CarpetaObservar, ArchivoBuscarAyer);
                    if (archivos.Length == 0)
                    {//si no encontró con fecha de ayer los buscamos con fecha de hoy 
                        archivos = Directory.GetFiles(CarpetaObservar, ArchivoBuscarHoy);                                            
                    }
                    else if (archivos.Length > 1)
                    { //hay más de un archivo para el día de ayer, nos quedamos únicamente con el más reciente, es decir con el que fue generado más tarde en el día. 
                        Utilerias.WriteToLog("Hay más de un archivo para : " + ArchivoBuscarAyer, "EnviaArchivos", Application.StartupPath + "\\Log.txt");
                        archivos = RegresaElUltimo(archivos);
                        Utilerias.WriteToLog("El archivo generado más reciente: " + archivos[0].Trim() + " del diá de ayer", "EnviaArchivos", Application.StartupPath + "\\Log.txt");
                    }

                    if (archivos.Length > 0)
                    {
                        foreach (string archivo in archivos)
                            {
                                FileInfo archtomove = new FileInfo(archivo);
                                string aux = "CierreDia_Conciliacion_Modulos vs contabilidad_13112013";
                                string nuevonombre = NombresMarcas[indicemarca].Trim() + "_" + strFechaHoy.Trim() + ".pdf";
                                archtomove.CopyTo(this.DirLocalArchivosZipear + "\\" + nuevonombre.Trim(),true);
                                Utilerias.WriteToLog("Se copia y renombra el archivo: " + nuevonombre + " a la carpeta: " + this.DirLocalArchivosZipear, "EnviaArchivos", Application.StartupPath + "\\Log.txt");
                                i++;
                            }                                        
                     }                    
                    else{                         
                        Utilerias.WriteToLog("La carpeta: " + CarpetaObservar + " No contiene el archivo: " + ArchivoBuscarAyer, "EnviaArchivos", Application.StartupPath + "\\Log.txt");
                        Utilerias.WriteToLog("La carpeta: " + CarpetaObservar + " No contiene el archivo: " + ArchivoBuscarHoy, "EnviaArchivos", Application.StartupPath + "\\Log.txt");
                    }
                    indicemarca++;
                }

                //20131216 Nos traemos todos los archivos .pdf que esten en la carpeta designada en el servidor remoto, los ponemos en la carpeta que se va a zipear.
                if (this.EnviaCorreoPrincipal.Trim() == "SI")
                {
                    i += TraeArchivosDesdeCarpetaRemota(this.CarpetaRemota, "*.*");
                }

                if (i > 0 && this.EnviaCorreoPrincipal.Trim()=="SI")
                { 
                  //Zippeamos y enviamos                                    
                  string ArchZip = Utilerias.ComprimirZip(this.DirLocalArchivosZipear, strFechaHoy + ".zip", "*.*");
                  if (ArchZip.Trim() != "")
                  {
                      Utilerias.WriteToLog("Se creo el zip:" + ArchZip.Trim() + " intentamos enviarlo", "EnviaArchivos:Antes enviar Correo", Application.StartupPath + "\\Log.txt");
                      //Enviamos el correo.
                      string rutaplantilla = Application.StartupPath;
                      rutaplantilla += "\\PlantillaGenOk.html";
                      string rutalogo = Application.StartupPath;
                      rutalogo += "\\image2993.png";

                      clsEmail correo = new clsEmail(this.smtpserverhost.Trim(), Convert.ToInt16(this.smtpport), this.usrcredential.Trim(), this.usrpassword.Trim(),this.EnableSsl.Trim());
                      MailMessage mnsj = new MailMessage();
                      mnsj.BodyEncoding = System.Text.Encoding.UTF8;
                      mnsj.Priority = System.Net.Mail.MailPriority.Normal;
                      mnsj.IsBodyHtml = true;
                      mnsj.Subject = this.PreSubject.Trim() + " "  + "Cierre Diario Conciliacion Modulos vs Contabilidad de la fecha: " + strFechaHoy;
                      string Remitente = "Sistemas de Grupo Andrade";

                      string[] Emails = this.emailsavisar.Split(';');
                      foreach (string Email in Emails)
                      {
                          mnsj.To.Add(new MailAddress(Email.Trim()));
                      }

                      mnsj.From = new MailAddress(usrcredential.Trim(), Remitente.Trim());
                      mnsj.Attachments.Add(new Attachment(ArchZip));

                      Dictionary<string, string> TextoIncluir = new Dictionary<string, string>();

                      TextoIncluir.Add("fecha", DateTime.Now.ToString("dd-MM-yyyy"));
                      TextoIncluir.Add("hora", DateTime.Now.ToString("HH:mm:ss"));
                      TextoIncluir.Add("fechaayer", strFechaHoy);
                      TextoIncluir.Add("iplocal", this.IPLocal);

                      #region cuando se desee que el usuario final sepa cuales agencias no se les generó el reporte
                      /*
                      if (arrAgenciasSinArchivo.Trim() != "")
                      {
                          string Renglon = "<tr><td align='left' valign='middle' bgcolor='#ffffff' height='29><p style='padding: 12px 0 0 17px; font-size: 10px; font-family:Trebuchet MS, Verdana, Arial, Helvetica, sans-serif; color:#000000; font-weight: bold; line-height:0;'>[Texto]</p></td></tr>";
                          string cadHTML = "";
                          
                          string[] arrAgencias = arrAgenciasSinArchivo.Split(',');
                          foreach (string Agenciasin in arrAgencias)
                          {
                              if (Agenciasin.Trim() != "")
                              {
                                  cadHTML += Renglon.Replace("[Texto]", Agenciasin.Trim()); 
                              }                          
                          }
                          
                          TextoIncluir.Add("TituloAgenciasSinArchivo", "Las siguientes agencias no se les generó reporte: ");
                          TextoIncluir.Add("AgenciasSinArchivo", cadHTML.Trim());
                      }
                      else{
                          TextoIncluir.Add("TituloAgenciasSinArchivo", "&nbsp;");
                          TextoIncluir.Add("AgenciasSinArchivo", " ");                      
                      }*/
                      #endregion
                      
                      TextoIncluir.Add("TituloAgenciasSinArchivo", "&nbsp;");
                      TextoIncluir.Add("AgenciasSinArchivo", " ");                      

                      AlternateView vistaplana = AlternateView.CreateAlternateViewFromString(correo.CreaCuerpoPlano(TextoIncluir), null, "text/plain");
                      AlternateView vistahtml = AlternateView.CreateAlternateViewFromString(correo.CreaCuerpoHTML(rutaplantilla, TextoIncluir).ToString(), null, "text/html");
                      LinkedResource logo = new LinkedResource(rutalogo);
                      logo.ContentId = "companylogo";
                      vistahtml.LinkedResources.Add(logo);

                      mnsj.AlternateViews.Add(vistaplana);
                      mnsj.AlternateViews.Add(vistahtml);

                      correo.MandarCorreo(mnsj);

                      res = 1;
                  }
                }
                    
                     string arrAgenciasSinArchivo = ValidaLaYaExistencia();
                     Utilerias.WriteToLog("Listado de agencias que no tienen archivo: " + arrAgenciasSinArchivo.Trim(), "Envia_Load", Application.StartupPath + "\\Log.txt");
                     //si hay agencias que no se les generó el reporte debemos informarlo, únicamente a la gente de sistemas                                                           
                     if (arrAgenciasSinArchivo.Trim() != "" && this.EnviaCorreoPrincipal.Trim()=="SI")
                      {                       
                      string rutaplantilla = Application.StartupPath;
                      rutaplantilla += "\\PlantillaAvisoNoSeGeneraron.txt";
                      clsEmail correoNoSeGenero = new clsEmail(this.smtpserverhost.Trim(), Convert.ToInt16(this.smtpport), this.usrcredential.Trim(), this.usrpassword.Trim(),this.EnableSsl.Trim());
                      MailMessage mensaje = new MailMessage();
                      mensaje.Priority = System.Net.Mail.MailPriority.Normal;
                      mensaje.IsBodyHtml = false;
                      mensaje.Subject = this.PreSubject.Trim() + " " + "Reportes No generados del Cierre Diario Conc. M vs Contabilidad de la fecha: " + strFechaHoy;
                      string Remitente = "Sistemas de Grupo Andrade";

                      string[] EmailsEspeciales = this.emailsavisarNoHayNada.Split(',');

                      foreach (string Email in EmailsEspeciales)
                      {
                          mensaje.To.Add(new MailAddress(Email.Trim()));
                      }

                      mensaje.From = new MailAddress(usrcredential.Trim(), Remitente.Trim());
                      //mnsj.Attachments.Add(new Attachment(ArchZip));

                      Dictionary<string, string> TextoIncluir = new Dictionary<string, string>();

                      TextoIncluir.Add("fecha", DateTime.Now.ToString("dd-MM-yyyy"));
                      TextoIncluir.Add("hora", DateTime.Now.ToString("HH:mm:ss"));
                      TextoIncluir.Add("iplocal", this.IPLocal);

                      TextoIncluir.Add("fechahoy", strFechaHoy);
                      arrAgenciasSinArchivo = arrAgenciasSinArchivo.Replace(",", "\n\r");
                      TextoIncluir.Add("agenciassinreporte", arrAgenciasSinArchivo);

                      //AlternateView vistaplana = AlternateView.CreateAlternateViewFromString(correo.CreaCuerpoPlano(TextoIncluir), null, "text/plain");
                      AlternateView vistahtml = AlternateView.CreateAlternateViewFromString(correoNoSeGenero.CreaCuerpoHTML(rutaplantilla, TextoIncluir).ToString(), null, "text/plain");

                      //LinkedResource logo = new LinkedResource(rutalogo);
                      //logo.ContentId = "companylogo";
                      //vistahtml.LinkedResources.Add(logo);

                      //mensaje.AlternateViews.Add(vistaplana);
                      mensaje.AlternateViews.Add(vistahtml);
                      correoNoSeGenero.MandarCorreo(mensaje);
                      res = 2;
                  }//cuando no se generó para alguna agencia.
                     
                res += EnviaArchivosXMarca();                    
            
            }
            catch (Exception ex)
            {
                Utilerias.WriteToLog("Error al enviar los archivos. \n" + ex.Message, "EnviaArchivos", Application.StartupPath + "\\Log.txt"); 
                Debug.WriteLine(ex.Message);
            }

            return res;
        }


        /// <summary>
        /// Para cada marca declarada en el dictionario, zippeará y enviará esos archivos a los correos definidos
        /// </summary>
        /// <returns>El número de correos enviados.</returns>
        public int EnviaArchivosXMarca()
        { 
            int res=0;

            try{

                DateTime FechaAyer = DateTime.Today.AddDays(-1);
                string FechaHoy = FechaAyer.ToString("ddMMyyyy");

                foreach (KeyValuePair<string, string> pair in this.dic_emailsxmarca)
                {
                    string NombreMarca = pair.Key.Trim();
                    string correos = pair.Value.ToString().Trim();
                    string patronbuscar = NombreMarca.Trim() + "*.pdf";
                    
                    //Zippeamos y enviamos                                    
                    string ArchZip = Utilerias.ComprimirZip(this.DirLocalArchivosZipear, NombreMarca + "_" + FechaHoy + ".zip",patronbuscar);
                    if (ArchZip.Trim() != "")
                    {
                        //Enviamos el correo.
                        string rutaplantilla = Application.StartupPath;
                        rutaplantilla += "\\PlantillaGenOk.html";
                        string rutalogo = Application.StartupPath;
                        rutalogo += "\\image2993.png";

                        clsEmail correo = new clsEmail(this.smtpserverhost.Trim(), Convert.ToInt16(this.smtpport), this.usrcredential.Trim(), this.usrpassword.Trim(),this.EnableSsl.Trim());
                        MailMessage mnsj = new MailMessage();
                        mnsj.BodyEncoding = System.Text.Encoding.UTF8;
                        mnsj.Priority = System.Net.Mail.MailPriority.Normal;
                        mnsj.IsBodyHtml = true;
                        mnsj.Subject = this.PreSubject.Trim() + " " + "Cierre Diario Conciliacion Modulos vs Contabilidad de la fecha: " + FechaHoy + "_" + NombreMarca.Trim();
                        string Remitente = "Sistemas de Grupo Andrade";

                        string[] Emails = correos.Split(',');
                        foreach (string Email in Emails)
                        {
                            mnsj.To.Add(new MailAddress(Email.Trim()));
                        }

                        mnsj.From = new MailAddress(usrcredential.Trim(), Remitente.Trim());
                        mnsj.Attachments.Add(new Attachment(ArchZip));

                        Dictionary<string, string> TextoIncluir = new Dictionary<string, string>();

                        TextoIncluir.Add("fecha", DateTime.Now.ToString("dd-MM-yyyy"));
                        TextoIncluir.Add("hora", DateTime.Now.ToString("HH:mm:ss"));
                        TextoIncluir.Add("fechaayer", FechaHoy);
                        TextoIncluir.Add("iplocal", this.IPLocal);

                        TextoIncluir.Add("TituloAgenciasSinArchivo", "&nbsp;");
                        TextoIncluir.Add("AgenciasSinArchivo", " ");

                        AlternateView vistaplana = AlternateView.CreateAlternateViewFromString(correo.CreaCuerpoPlano(TextoIncluir), null, "text/plain");
                        AlternateView vistahtml = AlternateView.CreateAlternateViewFromString(correo.CreaCuerpoHTML(rutaplantilla, TextoIncluir).ToString(), null, "text/html");
                        LinkedResource logo = new LinkedResource(rutalogo);
                        logo.ContentId = "companylogo";
                        vistahtml.LinkedResources.Add(logo);

                        mnsj.AlternateViews.Add(vistaplana);
                        mnsj.AlternateViews.Add(vistahtml);

                        correo.MandarCorreo(mnsj);

                        res++;
                    }
                    else {
                        Utilerias.WriteToLog("No hay archivos pdf para el patrón:" + patronbuscar.Trim(), "EnviaArchivosXMarca", Application.StartupPath + "\\Log.txt"); 
                    }
                }//del foreach
            }                                                  
            catch (Exception ex)
            {
                Utilerias.WriteToLog("Error al enviar los archivos. \n" + ex.Message, "EnviaArchivosXMarca", Application.StartupPath + "\\Log.txt"); 
                Debug.WriteLine(ex.Message);
            }

            return res;        

        }



        public int TraeArchivosDesdeCarpetaRemota(string DirRemoto,string filtro)
        {
            int res = 0;                       
            try
            {

                        #region funciones de logueo
                            IntPtr token = IntPtr.Zero;
                            IntPtr dupToken = IntPtr.Zero;

                            bool isSuccess = LogonUser(this.Usr, this.IPRemoto, this.Pass, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref token);
                            if (!isSuccess)
                            {
                                RaiseLastError();
                            }

                            isSuccess = DuplicateToken(token, 2, ref dupToken);
                            if (!isSuccess)
                            {
                                RaiseLastError();
                            }
                        #endregion
                            //una vez autenticado procedemos a traernos los archivos localizados en la carpeta remota.

                            WindowsIdentity newIdentity = new WindowsIdentity(dupToken);
                            using (newIdentity.Impersonate())
                            {
                                    try
                                    {
                                        DateTime hora_inicio = DateTime.Now;
                                        
                                         string[] arrfiles = Directory.GetFiles(DirRemoto, filtro);
                                         if (arrfiles.Length > 0)
                                         {
                                             foreach (string Arch in arrfiles)
                                             {
                                                 FileInfo Archivo = new FileInfo(Arch);
                                                 string nuevaruta = this.DirLocalArchivosZipear + "\\" + Archivo.Name.Trim();
                                                 Archivo.CopyTo(nuevaruta); //Esta linea es la que nos trae el archivo.
                                                 Utilerias.WriteToLog("Se trajo el archivo: " + nuevaruta  , "TraeArchivosDesdeCarpetaRemota", Application.StartupPath + "\\Log.txt");
                                                 res++;
                                             }                                         
                                         }

                                        #region Cálculo del tiempo transcurrido
                                        double minutostotales = 0;
                                        string TiempoTranscurrido = "";
                                        TimeSpan diferencia = new TimeSpan(0, 0, 0);
                                        diferencia = DateTime.Now - hora_inicio;
                                        minutostotales = diferencia.TotalMinutes;
                                        //this.minutostranscurridos = minutostotales;
                                        double segundostotales = diferencia.TotalSeconds;
                                        //.5 minutos = 30 segundos       //86400 segundos = un dia = 24 hrs * 60 min * 60 seg
                                        if (Utilerias.DaParteEntera((segundostotales / 86400)) > 0)
                                            TiempoTranscurrido = Convert.ToString(Utilerias.DaParteEntera((segundostotales / 86400))) + " d ";
                                        if (Utilerias.DaParteEntera(((segundostotales % 86400) / 3600)) > 0)
                                            TiempoTranscurrido += Convert.ToString(Utilerias.DaParteEntera((segundostotales % 86400) / 3600)) + " h ";
                                        if (Utilerias.DaParteEntera(((segundostotales % 86400) % 3600) / 60) > 0)
                                            TiempoTranscurrido += Convert.ToString(Utilerias.DaParteEntera(((segundostotales % 86400) % 3600) / 60)) + " m ";
                                        if (Utilerias.DaParteEntera(((segundostotales % 86400) % 3600) % 60) > 0)
                                            TiempoTranscurrido += Convert.ToString(Convert.ToInt16((((segundostotales % 86400) % 3600) % 60))) + " s ";
                                        #endregion
                                                                                
                                    }
                                    catch (Exception e)
                                    {
                                        Debug.WriteLine(e.Message);
                                        Utilerias.WriteToLog("FAILURE: \r" + e.Message + "\r", "TraeArchivosDesdeCarpetaRemota", Application.StartupPath + "\\Log.txt");
                                    }

                                isSuccess = CloseHandle(token);
                                if (!isSuccess)
                                {
                                    RaiseLastError();
                                }
                            }                        
                                            
            }
            catch (Exception ex)
            {
                Utilerias.WriteToLog("Error al traer los archivos en la carpeta remota \n\r" + ex.Message, "TraeArchivosDesdeCarpetaRemota", Application.StartupPath + "\\Log.txt");
                Debug.WriteLine(ex.Message);
            }

            return res;
        }


/************************************************************************************************************************/


        public string Dummy()
        {
            string res = "";
            try
            {

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }


            return res;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            #region Cálculo del tiempo transcurrido
            double minutostotales = 0;
            string TiempoTranscurrido = "";
            TimeSpan diferencia = new TimeSpan(0, 0, 0);
            diferencia = DateTime.Now - this.hora_inicio;
            minutostotales = diferencia.TotalMinutes;
            //this.minutostranscurridos = minutostotales;
            double segundostotales = diferencia.TotalSeconds;
            //.5 minutos = 30 segundos       //86400 segundos = un dia = 24 hrs * 60 min * 60 seg
            if (Utilerias.DaParteEntera((segundostotales / 86400)) > 0)
                TiempoTranscurrido = Convert.ToString(Utilerias.DaParteEntera((segundostotales / 86400))) + " d ";
            if (Utilerias.DaParteEntera(((segundostotales % 86400) / 3600)) > 0)
                TiempoTranscurrido += Convert.ToString(Utilerias.DaParteEntera((segundostotales % 86400) / 3600)) + " h ";
            if (Utilerias.DaParteEntera(((segundostotales % 86400) % 3600) / 60) > 0)
                TiempoTranscurrido += Convert.ToString(Utilerias.DaParteEntera(((segundostotales % 86400) % 3600) / 60)) + " m ";
            if (Utilerias.DaParteEntera(((segundostotales % 86400) % 3600) % 60) > 0)
                TiempoTranscurrido += Convert.ToString(Convert.ToInt16((((segundostotales % 86400) % 3600) % 60))) + " s ";
            #endregion
            Debug.WriteLine("Tiempo: " + TiempoTranscurrido);

            if (minutostotales > Convert.ToInt16(this.MinutosEjecucion) && this.MinutosEjecucion != "0")
            {
                Utilerias.WriteToLog("Se ha alcanzado el tiempo de ejecucion, el aplicativo se detendrá Tiempo Transcurrido: " + TiempoTranscurrido, "timer1_Tick", Application.StartupPath + "\\Log.txt");
                this.timer1.Stop();
                this.timer1.Enabled = false;
                this.Close();
                Application.Exit();
                return;
            }

            #region Forzar el cerrado de CachaYEnvia.
                       
                       int horacerrar = Convert.ToInt32(this.HoraCerrarAplicacion.Substring(0, this.HoraCerrarAplicacion.IndexOf(":")));
                       int minutocerrar = Convert.ToInt32(this.HoraCerrarAplicacion.Substring(this.HoraCerrarAplicacion.IndexOf(":") + 1));
                       DateTime momentocerrar = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, horacerrar, minutocerrar, 0);
                       
                       DateTime ahorita = DateTime.Now;                        

                       if (ahorita >= momentocerrar)
                       {
                        Utilerias.WriteToLog("Se ha alcanzado la hora de cierre de la aplicacion " + this.HoraCerrarAplicacion + " el aplicativo se cerrará ", "timer1_Tick", Application.StartupPath + "\\Log.txt");
                        this.timer1.Stop();
                        this.timer1.Enabled = false;
                        this.Close();
                        Application.Exit();
                        }

            #endregion


        }

        

        private bool FileReadyToRead(string filePath, int maxDuration)
        {
            int readAttempt = 0;
            while (readAttempt < maxDuration)
            {
                readAttempt++;
                try
                {
                    using (StreamReader stream = new StreamReader(filePath))
                    {
                        return true;
                    }
                }
                catch
                {
                    System.Threading.Thread.Sleep(60000);
                }
            }
            return false;
        }


        private void Envia_Paint(object sender, PaintEventArgs e)
        {
            this.ntiBalloon.Icon = this.Icon;
            this.ntiBalloon.Text = "ENVIO ARCHIVOS EXTRACCION BI";
            this.ntiBalloon.Visible = true;
            this.ntiBalloon.ShowBalloonTip(1, "INICIALIZADO", " En espera de la recepción de un archivo ", ToolTipIcon.Info);
            this.Hide();
            this.Visible = false;
        }

        private void Envia_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.ntiBalloon.Visible = false;
            this.ntiBalloon = null;
        }
    }
}
