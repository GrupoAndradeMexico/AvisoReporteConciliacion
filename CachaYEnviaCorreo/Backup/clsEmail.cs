﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Diagnostics;
using System.IO;

namespace CachaYEnvia
{    
    class clsEmail
    {
        SmtpClient server = null;


        public clsEmail(string smtpserverhost, int smtpport, string usrcredential , string usrpassword,string EnableSsl)
	    {
            server = new SmtpClient(smtpserverhost, smtpport);  
            server.UseDefaultCredentials = false;        
            server.Credentials = new System.Net.NetworkCredential(usrcredential, usrpassword);            
            server.EnableSsl = EnableSsl.Trim() == "true" ? true : false;
        }

        public void MandarCorreo(MailMessage mensaje)
        {
          server.Send(mensaje);
        }


       
        public StringBuilder CreaCuerpoHTML(string plantilla, Dictionary<string,string> TextoIncluir)
        {
            try
            {
                StringBuilder res = new StringBuilder(File.ReadAllText(plantilla));
                foreach (KeyValuePair<string, string> kvp in TextoIncluir)
                {
                    res.Replace("[" + kvp.Key.ToString() + "]", kvp.Value.ToString());
                }
                
                //Corregimos todos los acentos.
                //res.Replace("á", "&aacute;");
                //res.Replace("é", "&eacute;");
                //res.Replace("í", "&iacute;");
                //res.Replace("ó", "&oacute;");
                //res.Replace("ú", "&uacute;");
                //res.Replace("ñ", "&ntilde;");

                return res;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        public string CreaCuerpoPlano(Dictionary<string, string> TextoIncluir)
        {
            string res = "";
            res = "  \n" + "\r";            
            
            foreach (KeyValuePair<string, string> kvp in TextoIncluir)
            {
                res = res + res.Replace("[" + kvp.Key.ToString() + "]", kvp.Value.ToString()) + "\n" + "\r";
            }
                       
            res += "";

            //Corregimos todos los acentos.
            res = res.Replace("á", "&aacute;");
            res = res.Replace("é", "&eacute;");
            res = res.Replace("í", "&iacute;");
            res = res.Replace("ó", "&oacute;");
            res = res.Replace("ú", "&uacute;");
            res = res.Replace("ñ", "&ntilde;");

            return res;
        }
    }
}
