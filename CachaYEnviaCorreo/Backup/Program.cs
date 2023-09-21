using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;

namespace CachaYEnvia
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {

            bool createdNew;
            Mutex singleInstanceMutex = new Mutex(true, Process.GetCurrentProcess().ProcessName, out createdNew);

            try
            {
                if (createdNew)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new Envia(args));
                }
                else
                {
                    Utilerias.WriteToLog("El aplicativo ya se encuentra en memoria, se cierra esta instancia", "Main del sistema", Application.StartupPath + "\\Log.txt");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Utilerias.WriteToLog(ex.Message, "Main del sistema", Application.StartupPath + "\\Log.txt");
            }
            finally
            {
                singleInstanceMutex.Close();
            }                        
        }
    }
}
