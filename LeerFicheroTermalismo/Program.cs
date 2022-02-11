using Sentry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using  System.Configuration;

namespace LeerFicheroTermalismo
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {


            using (SentrySdk.Init(o =>
            {
                // Tells which project in Sentry to send events to:
                o.Dsn = "https://c58cf100305846caa0604a69c9839217@o400775.ingest.sentry.io/6199798";
                // When configuring for the first time, to see what the SDK is doing:
                o.Debug = true;
                // Set traces_sample_rate to 1.0 to capture 100% of transactions for performance monitoring.
                // We recommend adjusting this value in production.
                o.TracesSampleRate = 1.0;
                o.Environment = ConfigurationManager.AppSettings["Environment"];
            }))
            {
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);

         

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }

          
           
        }
    }
}
