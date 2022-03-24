using System;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using System.Collections;
using System.Data.OleDb;
using MySql.Data.MySqlClient;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Sentry;
using System.Net;

namespace LeerFicheroTermalismo
{


    public partial class Form1 : Form
    {

        BDsqlserver.functionsBD misfuncionesBD;
        BDsqlserver.functionsBD misfuncionesBDExpedientes;

        MySqlConnection msqlConnection;
        MySqlCommand msqlCommand;
        MySqlCommand msqlCommandContact;


        string lectura_fichero = "";
        string lectura_fichero_log = "";
        string lectura_fichero_nif = "";
        int contador_lineas_fichero = 0;
        int contador_lineas_fichero_dos = 0;
        int numero_reservas_insertadas = 0;
        ArrayList nif_fichero_imserso_mes = new ArrayList();
        DateTime fechaactual;
        string fecha_actual_strg = "";
        string fichero_log;
        string fichero_advertencias;
        string fichero_advertenciasclientes;
        string fichero_advertenciascliente_ai;
        string fichero_PreReservasPendientesReserva;
        string fichero_PreReservasPendientesReserva_Cdodigo_actulizar;
        string fichero_repetidas_reservas;
        string fichero_advertencias_log;
        string contrato_balneario = "";
        string conexion_balneario = "";
        string mes_del_fichero = "";
        int contador_lineas_correctas_total = 0;
        string epr_GUID_tv = "";
        int plazas_concedidas = 0;
        string institucion = "Imserso";
        string advertencia = "";
        string destino_FINAL = "";

        string id_balneario_seleccionado = "";


        int mes_global = 0;

        DataTable mytableAI = new DataTable();

        string agencia_balneario = "";
        string tipo_habitacion_balneario = "";
        string edificio_aci = "";
        string formato_datetime = "";
        string SPE_GUID = "";
        string PEN_GUID = "";
        string ORR_GUID;
        string EDI_GUID;
        string EDI_GUID_default;
        string admin_balneario = "";

        ArrayList PrereservasInsertadas = new ArrayList();
        string CAMPOS_excel_fichero = "";
        string curFile = "";

        string nacion = "";

        bool ojos_linea = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, System.EventArgs e)
        {






     

            this.Top = (Screen.PrimaryScreen.WorkingArea.Height - this.Height) / 2;
            this.Left = (Screen.PrimaryScreen.WorkingArea.Width - this.Width) / 2;

            DateTime fechaactual = DateTime.Now;
            fecha_actual_strg = fechaactual.ToString("ddMMyyyy HHmm");



            misfuncionesBDExpedientes = new BDsqlserver.functionsBD(new Utils().Decrypt(ConfigurationManager.AppSettings["ConnectionString"],"balneario"));
            llenar_combo_balnearios();

            DataColumn col1 = new DataColumn("fecha_entrada");
            DataColumn col2 = new DataColumn("advertencia");
            col1.DataType = System.Type.GetType("System.DateTime");
            col2.DataType = System.Type.GetType("System.String");
            mytableAI.Columns.Add(col1);
            mytableAI.Columns.Add(col2);

          

        }


        public int LeerFilasExcel(string name_fichero)
        {





            string[] tempArray = textBox3.Lines;

            OleDbConnection conexion = null;
            DataSet dataSet = null;
            OleDbDataAdapter dataAdapter = null;
            string consultaHojaExcel = "Select * from [Sheet1]";
            int counter = 0;
            int contador_lineas_fichero = 0;

            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;

            foreach (String fichero in tempArray)
            {
                if (fichero.Length > 2)
                {


                    //Conectamos con el fichero excel para contar sus fila
                    string cadenaConexionArchivoExcel = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fichero + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';";
                    conexion = new OleDbConnection(cadenaConexionArchivoExcel);//creamos la conexion con la hoja de excel
                    conexion.Open(); //abrimos la conexion
                    //conexion.Open(); //abrimos la conexion




                    dt = conexion.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);



                    String[] excelSheets = new String[dt.Rows.Count];
                    int i = 0;

                    // Add the sheet name to the string array.
                    foreach (DataRow row in dt.Rows)
                    {
                        excelSheets[i] = row["TABLE_NAME"].ToString();
                        i++;
                    }

                    //MessageBox.Show(excelSheets[0]);

                    consultaHojaExcel = "Select * from [" + excelSheets[0] + "]";


                    dataAdapter = new OleDbDataAdapter(consultaHojaExcel, conexion); //traemos los datos de la hoja y las guardamos en un dataSdapter

                    dataSet = new DataSet(); // creamos la instancia del objeto DataSet
                    dataAdapter.Fill(dataSet, excelSheets[0].Replace("$", ""));

                    foreach (DataRow dr in dataSet.Tables[0].Rows)
                    {


                        contador_lineas_fichero_dos++;



                    }




                }



            }


            return contador_lineas_fichero;

        }

        [Obsolete]
        public void LeerDatosExcel(string fichero)
        {



            int counter = 0;
            int contador_ficheros = 0;
            int contador_lineas_correctas = 0;
            string line;
            char[] delimiters = new char[] { '\t' };
            string[] parts = new string[34];


            int IdSolicitanteAci = 0;
            int IdAcompAci = 0;
            string nifformateadosolictante = "";
            string nifformateadoconyuge = "";

            string[] tempArray = textBox3.Lines;
            bool InsertoPrereserva = false;
            int numerodenifs_mes = 0;
            string mes_fichero = "";
            string year_fichero = "";
            DateTime fechaactual = DateTime.Now;
            int contador_inicio_agencia = 0;

            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;




            OleDbConnection conexion = null;
            DataSet dataSet = null;
            OleDbDataAdapter dataAdapter = null;
            string consultaHojaExcel = "Select * from [hoja1$]";

            string cadenaConexionArchivoExcel = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fichero + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';";
            conexion = new OleDbConnection(cadenaConexionArchivoExcel);//creamos la conexion con la hoja de excel
            conexion.Open(); //abrimos la conexion




            dt = conexion.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);



            String[] excelSheets = new String[dt.Rows.Count];
            int i = 0;

            // Add the sheet name to the string array.
            foreach (DataRow row in dt.Rows)
            {
                excelSheets[i] = row["TABLE_NAME"].ToString();
                i++;
            }

            //MessageBox.Show(excelSheets[0]);

            consultaHojaExcel = "Select * from [" + excelSheets[0] + "]";





            dataAdapter = new OleDbDataAdapter(consultaHojaExcel, conexion); //traemos los datos de la hoja y las guardamos en un dataSdapter

            dataSet = new DataSet(); // creamos la instancia del objeto DataSet
            dataAdapter.Fill(dataSet, excelSheets[0].Replace("$", ""));

            foreach (DataRow dr in dataSet.Tables[0].Rows)
            {



                IdAcompAci = 0;
                IdSolicitanteAci = 0;







                if (checkBox3.Checked == true)
                {
                    // MessageBox.Show(Convert.ToString(dr["NIF_SOLI"]));
                    parts[0] = Convert.ToString(DateTime.Now.Year + "/" + dr["EXPEDIENTE"]);
                    parts[1] = Convert.ToString(dr["NIF_SOLI"]);
                    parts[2] = Convert.ToString(dr["NOMBRE_SOLI"]);


                    this.advertencia = "";

                    if (Convert.ToString(dr["PLAZAS"]) == "2")
                    {


                        if (Convert.ToString(dr["NIF_CON"]) != "" && Convert.ToString(dr["NIF_CON"]).Length > 3)
                        {


                            parts[3] = Convert.ToString(dr["NIF_CON"]);
                            parts[4] = Convert.ToString(dr["NOMBRE_CON"]);
                            parts[31] = Convert.ToString(dr["FECHAN_CON"]);


                        }



                        if (Convert.ToString(dr["NIF_ACO"]) != "" && Convert.ToString(dr["NIF_ACO"]).Length > 3)
                        {


                            parts[3] = Convert.ToString(dr["NIF_ACO"]);
                            parts[4] = Convert.ToString(dr["NOMBRE_ACO"]);
                            parts[31] = Convert.ToString(dr["FECHAN_ACO"]);
                            this.advertencia = obtener_advertencia(id_balneario_seleccionado, "acom");

                        }



                        if (Convert.ToString(dr["NIF_HIJO"]) != "" && Convert.ToString(dr["NIF_HIJO"]).Length > 3)
                        {


                            parts[3] = Convert.ToString(dr["NIF_HIJO"]);
                            parts[4] = Convert.ToString(dr["NOMBRE_HIJO"]);
                            parts[31] = Convert.ToString(dr["FECHAN_HIJO"]);
                            this.advertencia = obtener_advertencia(id_balneario_seleccionado, "hijo");

                        }

                    }
                    else
                    {

                        parts[3] = "";
                        parts[4] = "";

                    }




                    if (Convert.ToString(dr["PLAZAS"]) == "2")
                    {

                        parts[5] = "C";


                    }

                    if (Convert.ToString(dr["PLAZAS"]) == "1")
                    {

                        parts[5] = "A";


                    }



                    parts[6] = "";
                    parts[7] = "31/12/" + DateTime.Now.Year;
                    parts[8] = Convert.ToString("A (7 DIAS)");
                    parts[9] = "";
                    parts[10] = "";
                    parts[11] = "";
                    parts[12] = Convert.ToString(dr["TELEFONO1"]);
                    parts[13] = Convert.ToString(dr["TELEFONO2"]);
                    parts[14] = Convert.ToString(dr["TEMPORADA"]);
                    parts[15] = "";
                    parts[16] = Convert.ToString(dr["Provincia"]);
                    parts[17] = Convert.ToString(dr["LOCALIDAD"]);
                    parts[18] = Convert.ToString(dr["CP"]);
                    parts[19] = Convert.ToString(dr["DIRECCION"]);



                    parts[30] = Convert.ToString(dr["FECHAN_SOLI"]);




                }
                else
                {



                    parts[0] = Convert.ToString(dr["Expediente"]);
                    parts[1] = Convert.ToString(dr["NIF Solicitante"]);
                    parts[2] = Convert.ToString(dr["Nombre S"]);
                    parts[3] = Convert.ToString(dr["NIF Conyuge"]);
                    parts[4] = Convert.ToString(dr["Nombre C"]);
                    parts[5] = Convert.ToString(dr["Plazas"]);
                    parts[6] = "";
                    parts[7] = Convert.ToString(dr["Fecha Turno"]);
                    parts[8] = Convert.ToString(dr["Duracion"]);
                    parts[9] = "";
                    parts[10] = "";
                    parts[11] = "";
                    parts[12] = Convert.ToString(dr["Telefono"]);
                    parts[13] = Convert.ToString(dr["Telefono2"]);
                    parts[14] = "";
                    parts[15] = "";
                    parts[16] = Convert.ToString(dr["Provincia"]);
                    parts[17] = Convert.ToString(dr["Localidad"]);
                    parts[18] = Convert.ToString(dr["cp"]);
                    parts[19] = Convert.ToString(dr["Direccion"]);

                    parts[30] = "";
                    parts[31] = "";




                }



                //MessageBox.Show(parts[0]);

                // MessageBox.Show(parts[1]);


                //MssageBox.Show(parts[7]);
                mes_fichero = Convert.ToString(Convert.ToDateTime(parts[7].Trim()).Month);
                year_fichero = Convert.ToString(Convert.ToDateTime(parts[7].Trim()).Year);
                this.mes_del_fichero = Convert.ToString(Convert.ToDateTime(parts[7].Trim()).Month);

                try
                {

                    if (this.checkBox3.Checked)
                    {
                        //curFile = textBox1.Text + @"\AdvertenciasTV_" + Convert.ToString(FechaImserso.Year) + "\\Advertencias_" + fecha_actual_strg + ".xls";
                        System.IO.Directory.CreateDirectory(textBox1.Text + @"\AdvertenciasTV_" + year_fichero);
                    }
                    else
                    {


                        //curFile = textBox1.Text + @"\AdvertenciasImserso" + mesfichero + "_" + anyofichero + "\\Advertencias_" + fecha_actual_strg + ".xls";
                        System.IO.Directory.CreateDirectory(textBox1.Text + @"\AdvertenciasImserso_" + mes_fichero + "_" + year_fichero);
                    }


                    // System.IO.Directory.CreateDirectory(textBox1.Text + @"\Advertencias" + mes_fichero + "_" + year_fichero);

                }
                catch
                {


                    //System.IO.Directory.Delete(textBox1.Text + @"\Advertencias\*.txt", true);

                }

                nifformateadosolictante = formatearnif(parts[1].Trim());
                nif_fichero_imserso_mes.Add(nifformateadosolictante);

                //MessageBox.Show(fechaactual.ToString("ddMMyyyy HHmm"));

                ojos_linea = false;
                fichero_advertencias += "###########################################################NifSol-" + parts[1].Trim() + "#################NifAcom-" + parts[3].Trim() + "#####################################################################################################################" + Environment.NewLine + Environment.NewLine;





                IdSolicitanteAci = ComprobarHuespedExiste(parts[1].Trim(), parts[2].Trim(), parts[19].Trim(), parts[17].Trim(), parts[18].Trim(), parts[16].Trim(), parts[12].Trim(), parts[13].Trim(), parts[7].Trim(), "Solicitante", nifformateadosolictante, parts[30].Trim(), parts[15].Trim());

                //Comprobamos si tiene acompañante
                if (Convert.ToString(parts[5].Trim()) == "C")
                {

                    this.plazas_concedidas = this.plazas_concedidas + 2;
                    nifformateadoconyuge = formatearnif(parts[3].Trim());
                    nif_fichero_imserso_mes.Add(nifformateadoconyuge);
                    IdAcompAci = ComprobarHuespedExiste(parts[3].Trim(), parts[4].Trim(), parts[19].Trim(), parts[17].Trim(), parts[18].Trim(), parts[16].Trim(), parts[12].Trim(), parts[13].Trim(), parts[7].Trim(), "Acompañante", nifformateadosolictante, parts[31].Trim(), parts[15].Trim());

                }
                else
                {


                    this.plazas_concedidas = this.plazas_concedidas + 1;


                }
                //Fin Comprobamos si tiene acompañante







                lectura_fichero_nif += IdSolicitanteAci + ";" + parts[0].Trim() + ";" + nifformateadosolictante + ";" + parts[2].Trim() + ";" + nifformateadoconyuge + ";" + parts[4].Trim() + ";" + parts[5].Trim() + ";" + parts[6].Trim() + ";" + parts[7].Trim() + ";" + parts[8].Trim() + ";" + parts[9].Trim() + ";" + parts[10].Trim() + ";" + parts[11].Trim() + ";" + parts[12].Trim() + ";" + parts[13].Trim() + ";" + parts[14].Trim() + ";" + parts[15].Trim() + parts[16].Trim() + ";" + parts[17].Trim() + ";" + parts[18].Trim() + System.Environment.NewLine;

                // MessageBox.Show(Convert.);


                //Comprvamos que no existe un prereserva para esa convocatoria tanto par el solicitente como el acompañante
                if ((ExistePrereserva(IdSolicitanteAci, parts, formatearnif(parts[1].Trim()), parts[1].Trim()) == 0) && (ExistePrereserva(IdAcompAci, parts, formatearnif(parts[3].Trim()), parts[3].Trim()) == 0))
                {


                    InsertoPrereserva = true;
                }
                else
                {


                    if (ConfigurationSettings.AppSettings["comprueba_reservas"] == "1")
                    {
                        InsertoPrereserva = false;
                    }
                    else
                    {

                        InsertoPrereserva = true;

                    }

                }





                //Comprobamos que no exite Reserva para esa convocatoria tanto par el solicitente como el acompañante
                if ((ExisteReserva(IdSolicitanteAci, parts, formatearnif(parts[1].Trim()), parts[1].Trim()) == 0) && (ExisteReserva(IdAcompAci, parts, formatearnif(parts[3].Trim()), parts[3].Trim()) == 0))
                {

                    if (InsertoPrereserva == true)
                    {

                        InsertoPrereserva = true;
                    }


                }
                else
                {


                    //Caso de que no quermos que comprobar si tien reserva

                    if (System.Configuration.ConfigurationSettings.AppSettings["comprueba_reservas"] == "0")
                    {


                        InsertoPrereserva = true;

                    }
                    else
                    {


                        InsertoPrereserva = false;

                    }




                }




                if (InsertoPrereserva == true)
                {


                    InsertarPrereserva(IdSolicitanteAci, IdAcompAci, parts, this.advertencia);

                }





                contador_lineas_correctas++;

                backgroundWorker2.ReportProgress(contador_lineas_correctas);
                //lectura_fichero += System.Environment.NewLine;



                fichero_advertencias += "################################################################################################################################################################################" + Environment.NewLine + Environment.NewLine;





                counter++;


            }

            //MessageBox.Show(Convert.ToString(counter));



            contador_ficheros++;



            //Comprueba los nifs de los huepedes
            HuespedesNoEstanTVExcel(mes_fichero, year_fichero, nif_fichero_imserso_mes);

            //inserto las advertencias
            //insetar_advertencias_fichero(mes_fichero, year_fichero);






        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {


            try
            {
                this.countLinesFile();
            }
            catch (Exception err)
            {
                SentrySdk.CaptureException(err);

            }



        }


        private void countLinesFile()
        {
            //Leer linea

            string[] tempArray = textBox3.Lines;
            int counter = 0;
            string line;
            //De cada fichero obtenemo sus datos



            //En el caso de que se un excel
            if (this.checkBox4.Checked == true || this.checkBox5.Checked == true)
            {




                if (this.checkBox4.Checked == true)
                {
                    try
                    {
                        LeerFilasExcel("Hoja1");
                    }
                    catch (Exception err)
                    {
                        SentrySdk.CaptureException(err);

                    }


                }





            }
            else
            {
                //En el caso de que se un excel




                foreach (String fichero in tempArray)
                {
                    if (fichero.Length > 2)
                    {
                        counter = 0;
                        System.IO.StreamReader file = new System.IO.StreamReader(fichero);
                        while ((line = file.ReadLine()) != null)
                        {
                            //MessageBox.Show(line);

                            if (this.checkBox3.Checked)
                            {
                                if (counter > 0)
                                {
                                    contador_lineas_fichero_dos++;
                                }
                            }
                            else
                            {

                                if (counter > 11)
                                {
                                    contador_lineas_fichero_dos++;

                                }


                            }



                            counter++;
                        }

                        //MessageBox.Show(Convert.ToString(contador_lineas_fichero_dos));


                    }


                    //MessageBox.Show(Convert.ToString(contador_lineas_fichero_dos));
                }









            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {

            this.progressBar1.Maximum = contador_lineas_fichero_dos;
            this.progressBar1.Minimum = 0;

            backgroundWorker2.RunWorkerAsync();


        }

        private void backgroundWorker2_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                this.readFile();
            }
            catch (Exception err)

            {
                
                SentrySdk.CaptureException(err);
                MessageBox.Show(err.ToString());

            }





        }

        private void readFile()
        {



            int counter = 0;
            int contador_ficheros = 0;
            int contador_lineas_correctas = 0;
            string line;
            char[] delimiters = new char[] { '\t' };
            string[] parts;
            int IdSolicitanteAci = 0;
            int IdAcompAci = 0;
            string nifformateadosolictante = "";
            string nifformateadoconyuge = "";

            string[] tempArray = textBox3.Lines;
            bool InsertoPrereserva = false;
            int numerodenifs_mes = 0;
            string mes_fichero = "";
            string year_fichero = "";
            DateTime fechaactual = DateTime.Now;
            int contador_inicio_agencia = 0;





            //De cada fichero obtenemo sus datos
            foreach (String fichero in tempArray)
            {



                counter = 0;
                IdAcompAci = 0;
                IdSolicitanteAci = 0;
                numerodenifs_mes = 0;
                nif_fichero_imserso_mes.Clear();
                mes_fichero = "";
                //Solo las lineas correctas



                if (fichero.Length > 2)
                {

                    fichero_log = "";
                    fichero_advertencias = "";
                    fichero_advertenciasclientes = "";

                    InsertoPrereserva = false;




                    // string[] ficherosCarpeta = Directory.GetFiles(textBox1.Text + @"\Advertencias" + mes_fichero + "_" + year_fichero);

                    if (this.checkBox3.Checked == true)
                    {

                        contador_inicio_agencia = 0;

                    }
                    else
                    {

                        contador_inicio_agencia = 11;
                    }




                    //En el cado de que sea un excel


                    //En el caso de que se un excel
                    if (this.checkBox4.Checked == true)
                    {


                        LeerDatosExcel(fichero);
                       


                        //Fin en el caso de que se un excel
                    }
                    else
                    {

                        //En el caso de que no sea un excel

                        // Read the file and display it line by line.
                        System.IO.StreamReader file = new System.IO.StreamReader(fichero);
                        while ((line = file.ReadLine()) != null)
                        {
                            // MessageBox.Show(  );
                            IdAcompAci = 0;
                            IdSolicitanteAci = 0;


                            //MessageBox.Show(Convert.ToString(counter)  +  Convert.ToString("--"+ contador_inicio_agencia) + "---"+line);

                            if (counter > contador_inicio_agencia)
                            {



                                parts = line.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                                //MessageBox.Show(Convert.ToString(counter) + "---" + Convert.ToString(parts.Length));
                                //MessageBox.Show(parts[0].Trim() + "--" + parts[1].Trim() + "--" + parts[2].Trim() + "--" + parts[3].Trim() + "--" + parts[4].Trim() + "--" + parts[5].Trim() + "--" + parts[6].Trim() + "--" + parts[7].Trim());




                                mes_fichero = Convert.ToString(Convert.ToDateTime(parts[7].Trim()).Month);
                                year_fichero = Convert.ToString(Convert.ToDateTime(parts[7].Trim()).Year);
                                this.mes_del_fichero = Convert.ToString(Convert.ToDateTime(parts[7].Trim()).Month);

                                try
                                {

                                    //System.IO.Directory.CreateDirectory(textBox1.Text + @"\Advertencias" + mes_fichero + "_" + year_fichero);

                                    if (this.checkBox3.Checked)
                                    {
                                        //curFile = textBox1.Text + @"\AdvertenciasTV_" + Convert.ToString(FechaImserso.Year) + "\\Advertencias_" + fecha_actual_strg + ".xls";
                                        System.IO.Directory.CreateDirectory(textBox1.Text + @"\AdvertenciasTV_" + year_fichero);
                                    }
                                    else
                                    {


                                        //curFile = textBox1.Text + @"\AdvertenciasImserso" + mesfichero + "_" + anyofichero + "\\Advertencias_" + fecha_actual_strg + ".xls";
                                        System.IO.Directory.CreateDirectory(textBox1.Text + @"\AdvertenciasImserso_" + mes_fichero + "_" + year_fichero);
                                    }



                                }
                                catch
                                {


                                    //System.IO.Directory.Delete(textBox1.Text + @"\Advertencias\*.txt", true);

                                }

                                nifformateadosolictante = formatearnif(parts[1].Trim());
                                nif_fichero_imserso_mes.Add(nifformateadosolictante);

                                //MessageBox.Show(fechaactual.ToString("ddMMyyyy HHmm"));

                                ojos_linea = false;
                                fichero_advertencias += "###########################################################NifSol-" + parts[1].Trim() + "#################NifAcom-" + parts[3].Trim() + "#####################################################################################################################" + Environment.NewLine + Environment.NewLine;

                                /* MessageBox.Show("13"+parts[13].Trim());
                                 MessageBox.Show("14" + parts[14].Trim());
                                 MessageBox.Show("15" + parts[15].Trim());
                                 MessageBox.Show("16" + parts[16].Trim());
                                 MessageBox.Show("17" + parts[17].Trim());
                                 MessageBox.Show("18" + parts[18].Trim());*/


                                IdSolicitanteAci = ComprobarHuespedExiste(parts[1].Trim(), parts[2].Trim(), parts[19].Trim(), parts[17].Trim(), parts[18].Trim(), parts[16].Trim(), parts[12].Trim(), parts[13].Trim(), parts[7].Trim(), "Solicitante", nifformateadosolictante, Convert.ToString(DateTime.Now), parts[15].Trim());

                                //Comprobamos si tiene acompañante
                                if (Convert.ToString(parts[5].Trim()) == "C")
                                {

                                    this.plazas_concedidas = this.plazas_concedidas + 2;
                                    nifformateadoconyuge = formatearnif(parts[3].Trim());
                                    nif_fichero_imserso_mes.Add(nifformateadoconyuge);
                                    IdAcompAci = ComprobarHuespedExiste(parts[3].Trim(), parts[4].Trim(), parts[19].Trim(), parts[17].Trim(), parts[18].Trim(), parts[16].Trim(), parts[12].Trim(), parts[13].Trim(), parts[7].Trim(), "Acompañante", nifformateadosolictante, Convert.ToString(DateTime.Now), parts[15].Trim());

                                }
                                else
                                {


                                    this.plazas_concedidas = this.plazas_concedidas + 1;


                                }
                                //Fin Comprobamos si tiene acompañante







                                lectura_fichero_nif += IdSolicitanteAci + ";" + parts[0].Trim() + ";" + nifformateadosolictante + ";" + parts[2].Trim() + ";" + nifformateadoconyuge + ";" + parts[4].Trim() + ";" + parts[5].Trim() + ";" + parts[6].Trim() + ";" + parts[7].Trim() + ";" + parts[8].Trim() + ";" + parts[9].Trim() + ";" + parts[10].Trim() + ";" + parts[11].Trim() + ";" + parts[12].Trim() + ";" + parts[13].Trim() + ";" + parts[14].Trim() + ";" + parts[15].Trim() + parts[16].Trim() + ";" + parts[17].Trim() + ";" + parts[18].Trim() + System.Environment.NewLine;


                                InsertoPrereserva = false;


                                //Comprobamos que no exite Reserva para esa convocatoria tanto par el solicitente como el acompañante
                                if ((ExisteReserva(IdSolicitanteAci, parts, formatearnif(parts[1].Trim()), parts[1].Trim()) == 0))
                                {

                                    //MessageBox.Show(Convert.ToString(parts[5].Trim()));

                                    if (parts[5].Trim() == "C")
                                    {

                                        //El caso de que vienn dos comprobamos el acompañante que no tenga reserva

                                        if ((ExisteReserva(IdAcompAci, parts, formatearnif(parts[3].Trim()), parts[3].Trim()) == 0))
                                        {

                                            InsertoPrereserva = true;



                                        }
                                        else
                                        {

                                            InsertoPrereserva = false;

                                        }


                                    }
                                    else
                                    {



                                        if (parts[5].Trim() == "B")
                                        {


                                            if ((ExisteReserva(IdAcompAci, parts, formatearnif(parts[3].Trim()), parts[3].Trim()) == 0))
                                            {

                                                InsertoPrereserva = true;
                                            }
                                            else
                                            {


                                                InsertoPrereserva = false;

                                            }


                                        }
                                        else
                                        {

                                            InsertoPrereserva = true;


                                        }





                                    }




                                }
                                else
                                {


                                    InsertoPrereserva = false;


                                }


                                //Comprobamos que no exite Reserva para esa convocatoria tanto par el solicitente como el acompañante
                                if ((ExistePrereserva(IdSolicitanteAci, parts, formatearnif(parts[1].Trim()), parts[1].Trim()) == 0) && InsertoPrereserva == true)
                                {

                                    if (parts[5].Trim() == "C")
                                    {

                                        //El caso de que vienn dos comprobamos el acompañante que no tenga reserva

                                        if ((ExistePrereserva(IdAcompAci, parts, formatearnif(parts[3].Trim()), parts[3].Trim()) == 0))
                                        {

                                            InsertoPrereserva = true;



                                        }
                                        else
                                        {

                                            InsertoPrereserva = false;

                                        }


                                    }
                                    else
                                    {


                                        if (parts[5].Trim() == "B")
                                        {


                                            if ((ExistePrereserva(IdAcompAci, parts, formatearnif(parts[3].Trim()), parts[3].Trim()) == 0))
                                            {

                                                InsertoPrereserva = true;
                                            }
                                            else
                                            {


                                                InsertoPrereserva = false;

                                            }


                                        }
                                        else
                                        {

                                            InsertoPrereserva = true;


                                        }

                                    }




                                }
                                else
                                {


                                    InsertoPrereserva = false;


                                }











                                //Todos los ok para intersan estn bien



                                if (InsertoPrereserva == true)
                                {



                                    // MessageBox.Show(Convert.ToString(InsertoPrereserva));
                                    InsertarPrereserva(IdSolicitanteAci, IdAcompAci, parts, "");

                                }





                                contador_lineas_correctas++;

                                backgroundWorker2.ReportProgress(contador_lineas_correctas);
                                //lectura_fichero += System.Environment.NewLine;



                                fichero_advertencias += "################################################################################################################################################################################" + Environment.NewLine + Environment.NewLine;



                            }

                            counter++;


                        }

                        //MessageBox.Show(Convert.ToString(counter));

                        file.Close();

                        contador_ficheros++;



                        //Comprueba los nifs de los huepedes

                        if (this.checkBox3.Checked)
                        {
                            HuespedesNoEstanTVExcel(mes_fichero, year_fichero, nif_fichero_imserso_mes);

                        }
                        else
                        {

                            HuespedesNoEstanImsersoExcel(mes_fichero, year_fichero, nif_fichero_imserso_mes);

                        }


                        //inserto las advertencias
                        //insetar_advertencias_fichero(mes_fichero, year_fichero);







                    }










                }











            }

        }

        private void backgroundWorker2_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {

            this.progressBar1.Value = e.ProgressPercentage;
            this.label1.Text = Convert.ToString(e.ProgressPercentage);
            this.label14.Text = Convert.ToString(this.plazas_concedidas);

            if (this.progressBar1.Maximum == Convert.ToInt32(this.label1.Text))
            {

                this.label11.Visible = true;

            }
            // this.ResultBlock.Text = lectura_fichero;
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {

            this.ResultBlock.Text = lectura_fichero;
            this.ResultBlock.Text += "#####################################################################################################################################" + System.Environment.NewLine;
            this.ResultBlock.Text += System.Environment.NewLine;
            this.ResultBlock.Text += "Numero de  Prereservas insetertadas " + Convert.ToString(numero_reservas_insertadas);
            this.ResultBlock.Text += System.Environment.NewLine;
            this.ResultBlock.Text += "#####################################################################################################################################" + System.Environment.NewLine;

            fichero_log += "#####################################################################################################################################" + System.Environment.NewLine;
            fichero_log += System.Environment.NewLine;
            fichero_log += "Numero de  Prereservas insetertadas " + Convert.ToString(numero_reservas_insertadas);
            fichero_log += System.Environment.NewLine;
            fichero_log += "#####################################################################################################################################" + System.Environment.NewLine;

            MessageBox.Show("Proceso Completado --> Numero de Prereservas insertadas --->" + Convert.ToString(numero_reservas_insertadas));
            this.label11.Visible = false;
            this.comboBox1.Enabled = true;

            this.button1.Enabled = true;
            this.button2.Enabled = true;
            this.button3.Enabled = true;
            this.textBox1.Enabled = true;
            this.textBox3.Enabled = true;
            this.ResultBlock.Enabled = true;
            this.dataGridView1.Enabled = true;
            this.dataGridView2.Enabled = true;
            this.dataGridView3.Enabled = true;
            this.textBox2.Enabled = true;
            this.button4.Enabled = true;


            // MessageBox.Show(this.mes_del_fichero);
            if (PrereservasInsertadas.Count > 0)
            {
                PrereservasPendientesIntersatadas();
            }
            ObtenerPrereservasPendientesRepetidas(this.mes_del_fichero);
            ObtenerPrereservasPendientes(this.mes_del_fichero);



            //System.IO.File.WriteAllText(@"C:\FicheroSalidaFomateado.txt", lectura_fichero_nif);
        }





        private void button1_Click(object sender, EventArgs e)
        {
            //openFileDialog1.Filter = "Text Files (.txt)|*.txt|All Files (*.*)|*.*";
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {

                textBox1.Text = this.folderBrowserDialog1.SelectedPath;
                //openFileDialog1.LoadFile(openFileDialog1.FileName, RichTextBoxStreamType.RichText);
            }

        }




        
        public int ComprobarHuespedExiste(string nif, string hue_des_str, string hue_dom_str, string hue_pob_str, string hue_cp_str, string hue_prv_str, string telefono1, string telefono2, string fecha_turno, string tipo_nif, string nifsolictente, string fecha_nac_d, string mail)
        {



            DataTable tabla_existe_registro = new DataTable();
            int HUE_GUID = 0;
            string nifformateado = "";

            /*Añadimos ceros en el caso de que el nif sea menor de nueve digitos*/
            nifformateado = formatearnif(nif);



            /*Buscamos el hue_guid si existe*/
            tabla_existe_registro = this.misfuncionesBD.obtenerDatable("SELECT HUE_GUID FROM [HuespedK] WHERE hue_nif_str='" + nifformateado + "' AND hue_nif_str IS NOT NULL AND hue_nif_str<> ''");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {
                HUE_GUID = Convert.ToInt32(dr["HUE_GUID"]);
            }
            /*FIn Buscamos el hue_guid si existe*/



            /*En el caso de que no existe el huesped lo creamos*/
            if (HUE_GUID == 0)
            {

                /*Insertamos el huesped*/
                HUE_GUID = InsertarHuespeded(hue_des_str.Trim(), hue_dom_str.Trim(), hue_pob_str.Trim(), hue_cp_str.Trim(), nifformateado, hue_prv_str.Trim(), telefono1, telefono2, fecha_turno, tipo_nif, nifsolictente, fecha_nac_d, mail);

            }
            else
            {
                //Actulizamos los campos necesarios que falten del huesped

                ActulizarHuespeded(HUE_GUID, hue_des_str.Trim(), hue_dom_str.Trim(), hue_pob_str.Trim(), hue_cp_str.Trim(), nifformateado, hue_prv_str.Trim(), telefono1, telefono2, fecha_turno, tipo_nif, nifsolictente, fecha_nac_d);
            }






            //Devolvemos el identificador de huesped
            return HUE_GUID;


        }



      

        public int InsertarHuespeded(string hue_des_str, string hue_dom_str, string hue_pob_str, string hue_cp_str, string hue_nif_str, string hue_prv_str, string telefono1, string telefono2, string fecha_turno, string tipo_nif, string nifsolictente, string fecha_nac, string email)
        {

            int HUE_GUID = 0;
            int Con_cod_lng = 0;
            string campos = "";
            string values = "";
            DateTime fecha_acutal = DateTime.Now;
            DateTime fecha_bird = Convert.ToDateTime(fecha_nac);
            string fecha_nacimiento = "";


            if (fecha_nac != "")
            {

                if (ValidateBirthday(fecha_nac))
                {

                    fecha_nacimiento = fecha_bird.ToString(formato_datetime);

                }
                else
                {

                    fecha_nacimiento = "";


                }



            }
            else
            {


                fecha_nacimiento = "";

            }

            Con_cod_lng = ObtnenerUltimoIdHuesped();
            HUE_GUID = ObtnenerUltimoHUE_GUIDHuesped();

            string nombre_apellidos = "";



            if (this.checkBox3.Checked)
            {


                char[] delimiterChars = { ' ' };
                string[] words_apellidos = hue_des_str.Split(delimiterChars);

                try
                {

                    switch (words_apellidos.Count())
                    {
                        case 3:
                            nombre_apellidos = words_apellidos[2] + " " + words_apellidos[1] + "," + words_apellidos[0];

                            break;
                        case 4:
                            nombre_apellidos = words_apellidos[3] + " " + words_apellidos[2] + "," + words_apellidos[0] + " " + words_apellidos[1];

                            break;
                        default:
                            nombre_apellidos = words_apellidos[2] + " " + words_apellidos[1] + "," + words_apellidos[0];

                            break;
                    }


                }
                catch
                {

                    nombre_apellidos = "";

                }



            }
            else
            {

                nombre_apellidos = hue_des_str.Replace("'", "''");


            }

            nombre_apellidos = nombre_apellidos;
            //insertamos en la table huespedes

            if (nombre_apellidos.Length > 40)
            {
                //MessageBox.Show(nombre_apellidos);

            }

            //MessageBox.Show(System.Configuration.ConfigurationSettings.AppSettings["nacion_" + this.nacion]);

            try
            {

                campos = "HUE_GUID,HUE_COD_str,hue_des_str,NAC_GUID,DIV_GUID,IDI_GUID,hue_co1_str,hue_co2_str,hue_fac_dat";
                values = "'" + HUE_GUID + "','" + Con_cod_lng + "','" + nombre_apellidos.Replace("'", "''") + "','" + System.Configuration.ConfigurationSettings.AppSettings["nacion_" + this.nacion] + "','2','2','','','" + fecha_acutal.ToString(formato_datetime) + "'";
                this.misfuncionesBD.insert("Huespedes", campos, values);

            }
            catch
            {

                campos = "HUE_GUID,HUE_COD_str,hue_des_str,NAC_GUID,DIV_GUID,IDI_GUID,hue_co1_str,hue_co2_str,hue_fac_dat";
                values = "'" + HUE_GUID + "','" + Con_cod_lng + "','','" + System.Configuration.ConfigurationSettings.AppSettings["nacion_" + this.nacion] + "','2','2','','','" + fecha_acutal.ToString(formato_datetime) + "'";
                this.misfuncionesBD.insert("Huespedes", campos, values);

            }





            //insertamos en la table huespedk

            //campos = "HUE_GUID,hue_dom_str,hue_pob_str,hue_cp_str,hue_nif_str,hue_prv_str,hue_pul_str,hue_dni_fna_dat,hue_mai_str";
            //values = "'" + HUE_GUID + "','" + hue_dom_str.Replace("'", "''") + "','" + hue_pob_str.Replace("'", "''") + "','" + hue_cp_str + "','" + hue_nif_str + "','" + hue_prv_str + "','TranspasoImserso','" + fecha_nacimiento + "','" + email + "'";
            campos = "HUE_GUID,hue_dom_str,hue_pob_str,hue_cp_str,hue_nif_str,hue_prv_str,hue_pul_str,hue_mai_str";
            values = "'" + HUE_GUID + "','" + hue_dom_str.Replace("'", "''") + "','" + hue_pob_str.Replace("'", "''") + "','" + hue_cp_str + "','" + hue_nif_str + "','" + hue_prv_str + "','TranspasoImserso','" + email + "'";


            this.misfuncionesBD.insert("HuespedK", campos, values);


            //insertamos los telefenos del huesped

            campos = "HUE_GUID,TEL_GUID,tel_tip_int,tel_num_str";
            values = "'" + HUE_GUID + "','1','1','" + telefono1 + "'";
            this.misfuncionesBD.insert("HuespedKTelefonos", campos, values);

            campos = "HUE_GUID,TEL_GUID,tel_tip_int,tel_num_str";
            values = "'" + HUE_GUID + "','2','1','" + telefono2 + "'";
            this.misfuncionesBD.insert("HuespedKTelefonos", campos, values);

            /**/

            this.misfuncionesBD.update("Contadors", "Con_cod_lng='" + Convert.ToString(Con_cod_lng + 1) + "'", "CON_GUID=26");

            lectura_fichero += "###############################################################################################################" + System.Environment.NewLine;
            lectura_fichero += "Huesped Insertado " + HUE_GUID + " con Nif para  " + tipo_nif + " " + hue_nif_str + " para el solicitante " + nifsolictente + System.Environment.NewLine;

            lectura_fichero_log = "Huesped Insertado " + HUE_GUID + " con Nif para  " + tipo_nif + " " + hue_nif_str + " para el solicitante " + nifsolictente + Environment.NewLine + Environment.NewLine;
            fichero_log += lectura_fichero_log;

            //MessageBox.Show(Convert.ToString(Con_cod_lng));
            //MessageBox.Show(Convert.ToString(HUE_GUID));


            //MessageBox.Show(HUE_GUID + "--" + Con_cod_lng + "--"); 

            return HUE_GUID;


        }








        //Obtenemos el ultmimo id para insertar en la tabla de huespedes

        public int ObtnenerUltimoIdHuesped()
        {
            int Con_cod_lng = 0;
            DataTable tabla_existe_registro = new DataTable();

            /*Obtenemos el id del ultimo contador*/
            tabla_existe_registro = this.misfuncionesBD.obtenerDatable("SELECT [CON_GUID],[Con_cod_lng] FROM [Contadors] WHERE [CON_GUID]=26");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {
                Con_cod_lng = Convert.ToInt32(dr["Con_cod_lng"]) + 1;
            }


            //Miramos en en aci si exite este huespeded

            for (int i = 1; i <= 200000; i++)
            {

                if (this.misfuncionesBD.existeregistro("SELECT HUE_COD_str FROM [Huespedes] WHERE HUE_COD_str='" + Con_cod_lng + "'", "HUE_COD_str"))
                {
                    //MessageBox.Show(Convert.ToString(Con_cod_lng));
                    Con_cod_lng = Con_cod_lng + i;

                }
                else
                {

                    break;
                }

            }

            //MessageBox.Show("Final"+Convert.ToString(Con_cod_lng));

            return Con_cod_lng;

        }



        //Obtenemos el ultmimo HUE_GUID para insertar en la tabla de huespedes

        public int ObtnenerUltimoHUE_GUIDHuesped()
        {
            int HUE_GUID = 0;
            DataTable tabla_existe_registro = new DataTable();

            /*Obtenemos el id del ultimo contador*/
            tabla_existe_registro = this.misfuncionesBD.obtenerDatable("SELECT MAX(HUE_GUID) AS MAXIMO FROM [Huespedes]");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {
                HUE_GUID = Convert.ToInt32(dr["MAXIMO"]) + 1;
            }



            return HUE_GUID;

        }




        //insertamos la prereserva
        public void InsertarPrereserva(int HUE_GUID, int HUE_GUID_ACOMPA, string[] huesped, string advr)
        {

            string campos = "";
            string values = "";
            string RES_GUID = "";
            string RES_GUID_COD = "";
            string AGE;
            string sqlInsert = "";
            string contrato = "";

            int adultos = 0;

            char[] delimiters = new char[] { '/' };
            string[] ExpedienteAno = huesped[0].Trim().Split(delimiters);


            DateTime fecha_Salida = DateTime.Now;

            DateTime fecha_acutal = DateTime.Now;

            RES_GUID = ObtnenerUltimoRES_GUIDPrereserva();
            RES_GUID_COD = RES_GUID;



            if (RES_GUID.Length < 8)
            {

                for (int i = 1; i <= (8 - RES_GUID.Length); i++)
                {

                    RES_GUID_COD = "0" + RES_GUID_COD;
                }

            }





            //En funcion del tipo de turno asigno la fecha de salida



            // Insertamos la prereserva

            string origen = "";


            if (this.checkBox3.Checked)
            {

                /* origen = "TransGeneralitat";
                 huesped[7] = "31/12/" + DateTime.Now.Year;


                 if (huesped[14] == "A")
                 {

                     contrato = System.Configuration.ConfigurationSettings.AppSettings["contrato_alta_tv"];

                 }
                 else
                 {


                     contrato = System.Configuration.ConfigurationSettings.AppSettings["contrato_baja_tv"];



                 }*/



                contrato = ObtenerContrato(Convert.ToDateTime(huesped[7].Trim()), "A (7 DIAS)", this.contrato_balneario);

            }
            else
            {

                //Otenemos el contrato
                contrato = ObtenerContrato(Convert.ToDateTime(huesped[7].Trim()), huesped[8].Trim(), this.contrato_balneario);


            }




            switch (huesped[8].Trim())
            {
                case "B (12 DIAS)":

                    fecha_Salida = Convert.ToDateTime(huesped[7].Trim()).AddDays(11);

                    break;

                case "A (10 DIAS)":

                    fecha_Salida = Convert.ToDateTime(huesped[7].Trim()).AddDays(9);

                    break;

                case "A (7 DIAS)":

                    fecha_Salida = Convert.ToDateTime(huesped[7].Trim()).AddDays(7);

                    break;


            }



            switch (huesped[5].Trim())
            {
                case "C":

                    adultos = 2;

                    break;


                case "A":

                    adultos = 1;

                    break;

                case "B":

                    adultos = 1;

                    break;

            }







            string agencia = "";

            agencia = this.agencia_balneario;


            string codigo_contrato = "";


            if (!this.checkBox3.Checked)
            {


                codigo_contrato = obtener_codigo_contrato(contrato);

            }
            else
            {

                codigo_contrato = contrato;

            }



            try
            {

                campos = "RES_GUID,RES_COD_str,res_fec_dat,res_ent_dat,res_sal_dat,res_dir_bln,TUR_GUID,AGE_GUID,EMP_GUID,HUE_GUID,res_nom_str,DIV_GUID,CTA_GUID,res_pro_dia_int,res_pro_por_cur,TIH_GUID,EDI_GUID,PEN_GUID,SRC_GUID_ENT,SRC_GUID_SAL,res_DR_bon_str,res_DR_vue_str,res_DR_hll_str,res_DR_hsa_str,res_DR_obs_str,res_est_int,res_anu_bln,res_ctr_fac_bln,res_nsh_bln,res_ce_lng,res_cf_lng,res_md_int,res_DC_rec_cur,res_DC_cre_bln,res_pag_dir_bln,res_pag_por_cur,res_anu_obs_str,res_spe_adu_int,res_spe_nin_int,res_spe_cun_int,res_gra_adu_int,res_gra_nin_int,res_gra_cun_int,res_no_alo_bln,res_Mul_bln,res_if_bln,res_res_bln,res_sim_bln,res_dto_sng,PLA_GUID,res_imp_inc_bln,res_gar_bln,res_tip_rmf_byt,res_com_sng,res_nal_bln,res_blo_bln,res_hwe_byt,res_pro_lng,res_gen_guid,res_gen_cod_str,res_gen_hue_guid,res_dto_com_sng,res_Msem_byt,res_Mdca_byt,res_Mnse_byt,res_Many_int,EPR_GUID,prr_est_byt,PRR_GUID,res_icn_bln,res_com_ind_sng,res_imp_inc_com_bln,PLA_COM_GUID,res_fra_sep_bln,res_srv_fij_lng,res_pre_med_cur,res_mrc_pln_lng,Mon_guid,Res_hor_byt,Moa_guid,res_ear_boo_sng,res_sse_bln,res_sdt_sng,res_sco_sng,res_sea_sng,PAQ_PRO_GUID,CNL_GUID,res_stf_bln,res_wel_val_bln,res_sup_lng,res_taf_tpv_lng,res_HUE_Apli_Fid,TAG_GUID,res_tip_byt,res_gen_hot_guid,res_cyt_lng,res_lco_bln,res_est_min_int,res_yie_byt,res_col_bar_lng,res_promo_NoDto_bln,res_promo_NoCom_bln,res_promo_NoEB_bln,res_win_byt,res_tax_tur_bln,res_tax_tit_lng,RHA_GUID,res_ear_com_sng,res_sdt_com_sng,res_sea_com_sng,res_sse_com_bln,res_ori_byt,ERC_GUID,res_het_bln,MET_GUID,res_tas_inc_bln,res_cop_byt,res_dtp_byt,res_eap_byt,res_master_bln";

                values = "'" + RES_GUID + "','" + RES_GUID_COD + "','" + fecha_acutal.ToString(formato_datetime) + "','" + Convert.ToDateTime(huesped[7].Trim()).ToString(formato_datetime) + "','" + Convert.ToDateTime(fecha_Salida).ToString(formato_datetime) + "','0','0','" + agencia + "','0','" + Convert.ToString(HUE_GUID) + "','','2','" + codigo_contrato + "','0','0.0000','" + tipo_habitacion_balneario + "','" + this.edificio_aci + "','" + this.PEN_GUID + "','2','1','" + ExpedienteAno[1] + "','','','','" + origen + "','2','0','0','0','0','2','0','0.0000','2','0','0,0000','','" + Convert.ToString(adultos) + "','0','0','0','0','0','0','0','0','0','0','0','1','True','0','1','0','0','0','0','0','0','0','0','0','0','0','0','0','" + obtener_estado_bd("Pendiente") + "','0','0','0','0','True',NULL,'0',NULL,'0',NULL,NULL,NULL,NULL,'0','0','0','0','0',NULL,'0','0','1','0','0',NULL,NULL,'1',NULL,'0','0','0','0','0','0','0','0',NULL,'0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0'";

                sqlInsert = "INSERT INTO PresReservas (" + campos + ") VALUES (" + values + ");";

                this.misfuncionesBD.insert("PresReservas", campos, values);

            }
            catch
            {

                campos = "RES_GUID,RES_COD_str,res_fec_dat,res_ent_dat,res_sal_dat,res_dir_bln,TUR_GUID,AGE_GUID,EMP_GUID,HUE_GUID,res_nom_str,DIV_GUID,CTA_GUID,res_pro_dia_int,res_pro_por_cur,TIH_GUID,EDI_GUID,PEN_GUID,SRC_GUID_ENT,SRC_GUID_SAL,res_DR_bon_str,res_DR_vue_str,res_DR_hll_str,res_DR_hsa_str,res_DR_obs_str,res_est_int,res_anu_bln,res_ctr_fac_bln,res_nsh_bln,res_ce_lng,res_cf_lng,res_md_int,res_DC_rec_cur,res_DC_cre_bln,res_pag_dir_bln,res_pag_por_cur,res_anu_obs_str,res_spe_adu_int,res_spe_nin_int,res_spe_cun_int,res_gra_adu_int,res_gra_nin_int,res_gra_cun_int,res_no_alo_bln,res_Mul_bln,res_if_bln,res_res_bln,res_sim_bln,res_dto_sng,PLA_GUID,res_imp_inc_bln,res_gar_bln,res_tip_rmf_byt,res_com_sng,res_nal_bln,res_blo_bln,res_hwe_byt,res_pro_lng,res_gen_guid,res_gen_cod_str,res_gen_hue_guid,res_dto_com_sng,res_Msem_byt,res_Mdca_byt,res_Mnse_byt,res_Many_int,EPR_GUID,prr_est_byt,PRR_GUID,res_icn_bln,res_com_ind_sng,res_imp_inc_com_bln,PLA_COM_GUID,res_fra_sep_bln,res_srv_fij_lng,res_pre_med_cur,res_mrc_pln_lng,Mon_guid,Res_hor_byt,Moa_guid,res_ear_boo_sng,res_sse_bln,res_sdt_sng,res_sco_sng,res_sea_sng,PAQ_PRO_GUID,CNL_GUID,res_stf_bln,res_wel_val_bln,res_sup_lng,res_taf_tpv_lng,res_HUE_Apli_Fid,TAG_GUID,res_tip_byt,res_gen_hot_guid,res_cyt_lng,res_lco_bln,res_est_min_int,res_yie_byt,res_col_bar_lng,res_promo_NoDto_bln,res_promo_NoCom_bln,res_promo_NoEB_bln,res_win_byt,res_tax_tur_bln,res_tax_tit_lng,RHA_GUID,res_ear_com_sng,res_sdt_com_sng,res_sea_com_sng,res_sse_com_bln,res_ori_byt,ERC_GUID,res_het_bln,MET_GUID";

                values = "'" + RES_GUID + "','" + RES_GUID_COD + "','" + fecha_acutal.ToString(formato_datetime) + "','" + Convert.ToDateTime(huesped[7].Trim()).ToString(formato_datetime) + "','" + Convert.ToDateTime(fecha_Salida).ToString(formato_datetime) + "','0','0','" + agencia + "','0','" + Convert.ToString(HUE_GUID) + "','','2','" + codigo_contrato + "','0','0.0000','" + tipo_habitacion_balneario + "','" + this.edificio_aci + "','" + this.PEN_GUID + "','2','1','" + ExpedienteAno[1] + "','','','','" + origen + "','2','0','0','0','0','2','0','0.0000','2','0','0,0000','','" + Convert.ToString(adultos) + "','0','0','0','0','0','0','0','0','0','0','0','1','True','0','1','0','0','0','0','0','0','0','0','0','0','0','0','0','" + obtener_estado_bd("Pendiente") + "','0','0','0','0','True',NULL,'0',NULL,'0',NULL,NULL,NULL,NULL,'0','0','0','0','0',NULL,'0','0','1','0','0',NULL,NULL,'1',NULL,'0','0','0','0','0','0','0','0',NULL,'0','0','0','0','0','0','0','0','0','0','0'";

                sqlInsert = "INSERT INTO PresReservas (" + campos + ") VALUES (" + values + ");";

                this.misfuncionesBD.insert("PresReservas", campos, values);
            }



            //obtnemos todas las advertencias que tiene el cliente

            DataTable tabla_existe_advertenicias = new DataTable();


            // this.misfuncionesBD.delete("AdvertenciaAsignadaReservas"," AAR_GUID ='"+RES_GUID+"');
            /*Obtenemos el id del ultimo contador*/
            tabla_existe_advertenicias = this.misfuncionesBD.obtenerDatable("SELECT  AdvertenciaAsignadas.ADV_GUID, AdvertenciaAsignadas.AAS_GUID, AdvertenciaAsignadas.AAS_tab_str, AdvertenciaAsignadas.AAS_obs_str  FROM  AdvertenciaAsignadas INNER JOIN HuespedK ON AdvertenciaAsignadas.AAS_GUID = HuespedK.HUE_GUID AND AdvertenciaAsignadas.AAS_tab_str = 'Huespedes' WHERE HuespedK.HUE_GUID='" + Convert.ToString(HUE_GUID) + "'");

            foreach (DataRow dr in tabla_existe_advertenicias.Rows)
            {

                try
                {

                    this.misfuncionesBD.insert("AdvertenciaAsignadaReservas", "ADV_GUID,AAR_GUID,AAR_reg_bln,DEP_GUID,AAR_alk_bln,AAR_con_bln", "'" + Convert.ToInt32(dr["ADV_GUID"]) + "','" + RES_GUID + "','0','0','False','False'");

                }
                catch
                {


                }

                //HUE_GUID = Convert.ToInt32(dr["MAXIMO"]) + 1;
            }










            //Actulizo la reserva con el origen de esta
            if (this.checkBox3.Checked)
            {

                this.misfuncionesBD.update("PresReservas", "epr_GUID='" + this.epr_GUID_tv + "'", "RES_GUID=" + RES_GUID);
                this.misfuncionesBD.delete("AdvertenciaAsignadaReservas", " AAR_GUID='" + RES_GUID + "' AND ADV_GUID='" + advr + "'");
                this.misfuncionesBD.insert("AdvertenciaAsignadaReservas", "ADV_GUID,AAR_GUID,AAR_reg_bln,DEP_GUID,AAR_alk_bln,AAR_con_bln", "'" + advr + "','" + RES_GUID + "','0','0','False','False'");


            }





            this.misfuncionesBD.update("PresReservas", "ORR_GUID='" + this.ORR_GUID + "'", "RES_GUID=" + RES_GUID);



            //lectura_fichero += sqlInsert  + System.Environment.NewLine;



            //Borro si existe la prereserva en la tabla PresReservaCantidadesDesglosado
            this.misfuncionesBD.delete("PresReservaCantidadesDesglosado", "RES_GUID='" + RES_GUID + "'");




            /*

            //Insertamso en prereservas desglosado

            campos = "RES_GUID,SCA_GUID,SPE_GUID,rsc_num_int";

            values = "'" + RES_GUID + "','1','" + SPE_GUID + "','" + Convert.ToString(adultos) + "'";

            sqlInsert = "INSERT INTO PresReservaCantidadesDesglosado (" + campos + ") VALUES  (" + values + ");";

            this.misfuncionesBD.insert("PresReservaCantidadesDesglosado", campos, values);

            // lectura_fichero += sqlInsert  + System.Environment.NewLine;




            campos = "RES_GUID,SCA_GUID,SPE_GUID,rsc_num_int";

            values = "'" + RES_GUID + "','2','" + SPE_GUID + "','" + Convert.ToString(adultos) + "'";

            sqlInsert = "INSERT INTO PresReservaCantidadesDesglosado (" + campos + ") VALUES (" + values + ");";

            this.misfuncionesBD.insert("PresReservaCantidadesDesglosado", campos, values);

            //lectura_fichero += sqlInsert  + System.Environment.NewLine;


            campos = "RES_GUID,SCA_GUID,SPE_GUID,rsc_num_int";

            values = "'" + RES_GUID + "','4','" + SPE_GUID + "','" + Convert.ToString(adultos) + "'";

            sqlInsert = "INSERT INTO PresReservaCantidadesDesglosado (" + campos + ") VALUES  (" + values + ");";

            this.misfuncionesBD.insert("PresReservaCantidadesDesglosado", campos, values);

            //lectura_fichero += sqlInsert  + System.Environment.NewLine;

            */





            DataTable tabla_existe_registro_SPE_GUID = new DataTable();
            tabla_existe_registro_SPE_GUID = this.misfuncionesBDExpedientes.obtenerDatable("SELECT * FROM    agencias_SPE_GUID_SCA_GUID WHERE age_guid='" + agencia + "' AND destino ='" + this.destino_FINAL + "' ");
            int SPE_GUID_FINAL = 0;
            foreach (DataRow dr_SPE_GUID in tabla_existe_registro_SPE_GUID.Rows)
            {



                campos = "RES_GUID,SCA_GUID,SPE_GUID,rsc_num_int";

                values = "'" + RES_GUID + "','" + dr_SPE_GUID["SCA_GUID"] + "','" + dr_SPE_GUID["SPE_GUID"] + "','" + Convert.ToString(adultos) + "'";

                sqlInsert = "INSERT INTO PresReservaCantidadesDesglosado (" + campos + ") VALUES  (" + values + ");";

                this.misfuncionesBD.insert("PresReservaCantidadesDesglosado", campos, values);

                SPE_GUID_FINAL = Convert.ToInt32(dr_SPE_GUID["SPE_GUID"]);

            }





            //Insertamos el acompañante en la tabla de acompañantes

            if (adultos == 2)
            {

                //Borro si existe la prereserva en la tabla PresReservaAcompanantes
                this.misfuncionesBD.delete("PresReservaAcompanantes", "RES_GUID='" + RES_GUID + "'");

                campos = "RES_GUID,RAC_GUID,rac_nom_str,HUE_GUID,rac_tax_tur_bln,rac_het_bln,MET_GUID";

                values = "'" + RES_GUID + "','1','" + huesped[4].Trim() + "','" + Convert.ToString(HUE_GUID_ACOMPA) + "','0','0','0'";

                sqlInsert = "INSERT INTO PresReservaAcompanantes (" + campos + ") VALUES (" + values + ");";

                this.misfuncionesBD.insert("PresReservaAcompanantes", campos, values);

                //lectura_fichero += sqlInsert  + System.Environment.NewLine;


            }



            //Insertamos PresReservaPersonas

            //Borro si existe la prereserva en la tabla PresReservaPersonas
            this.misfuncionesBD.delete("PresReservaPersonas", "RES_GUID='" + RES_GUID + "'");

            campos = "RES_GUID,SPE_GUID,rde_num_int,rde_gra_int";

            values = "'" + RES_GUID + "','" + SPE_GUID_FINAL + "','" + Convert.ToString(adultos) + "','0'";

            sqlInsert = "INSERT INTO PresReservaPersonas (" + campos + ") VALUES  (" + values + ")";

            this.misfuncionesBD.insert("PresReservaPersonas", campos, values);


            PrereservasInsertadas.Add(RES_GUID);

            lectura_fichero += "###############################################################################################################" + System.Environment.NewLine;
            lectura_fichero += "PreReserva Insertda con codigo " + RES_GUID + " para el expediente " + ExpedienteAno[1] + " con Nif Solicntante " + formatearnif(huesped[1].Trim()) + " y con de Huesped en Aci " + Convert.ToString(HUE_GUID) + System.Environment.NewLine;
            lectura_fichero += "###############################################################################################################" + System.Environment.NewLine;

            fichero_log += "###############################################################################################################" + System.Environment.NewLine;
            fichero_log += "PreReserva Insertda con codigo " + RES_GUID + " para el expediente " + ExpedienteAno[1] + " con Nif Solicntante " + formatearnif(huesped[1].Trim()) + " y con de Huesped en Aci " + Convert.ToString(HUE_GUID) + System.Environment.NewLine;
            fichero_log += "###############################################################################################################" + System.Environment.NewLine;



            lectura_fichero_log = "PreReserva Insertada con codigo " + RES_GUID + " para el expediente " + ExpedienteAno[1] + " con Nif Solicntante " + formatearnif(huesped[1].Trim()) + " y con de Huesped en Aci " + Convert.ToString(HUE_GUID) + System.Environment.NewLine;
            numero_reservas_insertadas++;

            //InsentarLineaLog(Convert.ToDateTime(huesped[7].Trim()), lectura_fichero_log);



        }
        //Fin de insertar prereserva



        //Obtenemos el ultmimo RES_GUID para insertar en la tabla de prereservas

        public string ObtnenerUltimoRES_GUIDPrereserva()
        {
            int RES_GUID = 0;
            DataTable tabla_existe_registro = new DataTable();

            /*Obtenemos el id del ultimo contador*/
            tabla_existe_registro = this.misfuncionesBD.obtenerDatable("SELECT MAX(RES_GUID) AS MAXIMO FROM [PresReservas] ");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {
                RES_GUID = Convert.ToInt32(dr["MAXIMO"]) + 1;
            }


            if (RES_GUID == 0)
            {

                RES_GUID = 1;

            }


            return Convert.ToString(RES_GUID);

        }


        //Existe prereserva para esa convocatoria

        public int ExistePrereserva(int HUE_GUID, string[] huesped, string nif_huesped, string nif_original)
        {

            string nif = huesped[1].Trim();
            string rango_convocatorias = "";
            string numerodeadultos = "";
            string expedienteACi = "";
            DataSet miDataset = new DataSet();
            /*Añadimos ceros en el caso de que el nif sea menor de nueve digitos*/

            nif = formatearnif(nif);


            //Obtengo el rango de fechas de la convocatoria
            char[] delimiters = new char[] { '-' };
            rango_convocatorias = ObtenerRangoConvocatoria(Convert.ToDateTime(huesped[7].Trim()));
            string[] VectotConvocatorias = rango_convocatorias.Split(delimiters);


            char[] delimiters_expediente = new char[] { '/' };
            string[] VectorExpedienteAnyo = Convert.ToString(huesped[0].Trim()).Split(delimiters_expediente);


            int RES_GUID = 0;
            DataTable tabla_existe_registro = new DataTable();

            // lectura_fichero += System.Environment.NewLine + "SELECT RES_GUID FROM [PresReservas] WHERE HUE_GUID='" + Convert.ToString(HUE_GUID) + "' AND res_ent_dat between  convert(datetime, '" + Convert.ToDateTime(VectotConvocatorias[0]).ToString("dd/MM/yyyy") + "') AND  convert(datetime, '" + Convert.ToDateTime(VectotConvocatorias[1]).ToString("dd/MM/yyyy") + "')" + System.Environment.NewLine; 

            /*Obtenemos el id del ultimo contador*/
            string sql = "";



            if (this.checkBox3.Checked)
            {
                sql = "SELECT edi_des_str,res_nal_bln,age_nom_str,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[PresReservas].RES_GUID,HuespedK.HUE_GUID,PresReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,PreReservaEstados.epr_des_str as Estado,prr_est_byt FROM PresReservas INNER JOIN  PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID INNER JOIN HuespedK ON HuespedK.HUE_GUID = PresReservas.HUE_GUID LEFT  JOIN  PresReservaAcompanantes ON  PresReservaAcompanantes.RES_GUID = PresReservas.RES_GUID   INNER JOIN Agencias ON  Agencias.AGE_GUID = PresReservas.AGE_GUID  LEFT JOIN Edificios ON Edificios.EDI_GUID  = PresReservas.EDI_GUID  WHERE  PresReservas.AGE_GUID='" + this.agencia_balneario + "' AND HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND YEAR(res_ent_dat)=YEAR(GETDATE())    ";

            }
            else
            {

                sql = "SELECT edi_des_str,res_nal_bln,age_nom_str,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[PresReservas].RES_GUID,HuespedK.HUE_GUID,PresReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,PreReservaEstados.epr_des_str as Estado,prr_est_byt FROM PresReservas INNER JOIN  PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID INNER JOIN HuespedK ON HuespedK.HUE_GUID = PresReservas.HUE_GUID LEFT  JOIN  PresReservaAcompanantes ON  PresReservaAcompanantes.RES_GUID = PresReservas.RES_GUID   INNER JOIN Agencias ON  Agencias.AGE_GUID = PresReservas.AGE_GUID LEFT JOIN Edificios ON Edificios.EDI_GUID  = PresReservas.EDI_GUID   WHERE PresReservas.AGE_GUID='" + this.agencia_balneario + "' AND (HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND res_ent_dat between '" + Convert.ToDateTime(VectotConvocatorias[0]).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime(VectotConvocatorias[1]).ToString(formato_datetime) + "')  OR (res_ent_dat between '" + Convert.ToDateTime("01/01/" + DateTime.Now.Year).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime("31/12/" + DateTime.Now.Year).ToString(formato_datetime) + "' AND  res_DR_bon_str='" + VectorExpedienteAnyo[1] + "' ) ";


            }




            //MessageBox.Show(sql);

            //string sql_es = "SELECT res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[PresReservas].RES_GUID,HuespedK.HUE_GUID,PresReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,PreReservaEstados.epr_des_str as Estado,prr_est_byt FROM PresReservas INNER JOIN  PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID INNER JOIN HuespedK ON HuespedK.HUE_GUID = PresReservas.HUE_GUID LEFT  JOIN  PresReservaAcompanantes ON  PresReservaAcompanantes.RES_GUID = PresReservas.RES_GUID     WHERE (HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND res_ent_dat between '" + Convert.ToDateTime(VectotConvocatorias[0]).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime(VectotConvocatorias[1]).ToString(formato_datetime) + "') OR  (res_ent_dat between '01/01/" + Convert.ToDateTime(VectotConvocatorias[1]).Year + "' AND   '31/12/" + Convert.ToDateTime(VectotConvocatorias[1]).Year + "' AND  res_DR_bon_str='" + VectorExpedienteAnyo[1] + "' ) ";

            miDataset = this.misfuncionesBD.obtenerDataSet(sql);







            DateTime fechaReservaAci = DateTime.Now;
            string nif_huesped_aci = "";
            string IdAciHuesped = "";
            string estado = "";
            string idAcomp = "";
            string nif_acomp_Aci = "";
            string tipoturno = Convert.ToString(huesped[8].Trim());
            string diasestancia = "";
            char[] delimiters_re = new char[] { '/' };

            tabla_existe_registro = miDataset.Tables[0];
            string noalojados = "";
            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                RES_GUID = Convert.ToInt32(dr["RES_GUID"]);
                fechaReservaAci = Convert.ToDateTime(dr["res_ent_dat"]);
                nif_huesped_aci = formatearnif(Convert.ToString(dr["hue_nif_str"]));
                IdAciHuesped = Convert.ToString(dr["HUE_GUID"]);
                estado = Convert.ToString(dr["Estado"]);
                numerodeadultos = Convert.ToString(dr["res_spe_adu_int"]);
                idAcomp = Convert.ToString(dr["acomp"]);
                nif_acomp_Aci = ObtenerNifAcomAci(idAcomp);
                diasestancia = Convert.ToString(dr["DiferenciaDias"]);
                expedienteACi = Convert.ToString(dr["res_DR_bon_str"]);
                noalojados = Convert.ToString(dr["res_nal_bln"]);

                // InsentarPacientesExcelAdvertencias(Convert.ToDateTime(huesped[7].Trim()), fechaReservaAci, " Prereserva ", Convert.ToString(IdAciHuesped), Convert.ToString(nif), Convert.ToString(nif_huesped_aci), Convert.ToString(RES_GUID), VectorExpedienteAnyo[1], estado, numerodeadultos, Convert.ToString(huesped[5].Trim()), formatearnif(nif_acomp_Aci), formatearnif(huesped[3].Trim()), tipoturno, diasestancia, expedienteACi, "", Convert.ToString(dr["age_nom_str"]), VectorExpedienteAnyo[1]);
                if (!this.checkBox3.Checked)
                {

                    if (estado.ToUpper() == "NO")
                    {



                        InsentarPacientesExcelAdvertencias(Convert.ToDateTime(huesped[7].Trim()), fechaReservaAci, " Prereserva ", Convert.ToString(IdAciHuesped), Convert.ToString(huesped[1].Trim()), Convert.ToString(nif_huesped_aci), Convert.ToString(RES_GUID), VectorExpedienteAnyo[1], estado, numerodeadultos, Convert.ToString(huesped[5].Trim()), formatearnif(nif_acomp_Aci), formatearnif(huesped[3].Trim()), tipoturno, diasestancia, expedienteACi, noalojados, Convert.ToString(dr["age_nom_str"]), VectorExpedienteAnyo[1], HUE_GUID, huesped, nif_huesped, nif_original, Convert.ToString(dr["edi_des_str"]));


                    }
                }



                if (this.checkBox3.Checked)
                {
                    InsentarPacientesExcelAdvertenciasTV(Convert.ToDateTime(huesped[7].Trim()), fechaReservaAci, " Prereserva ", Convert.ToString(IdAciHuesped), Convert.ToString(huesped[1].Trim()), Convert.ToString(nif_huesped_aci), Convert.ToString(RES_GUID), VectorExpedienteAnyo[1], estado, numerodeadultos, Convert.ToString(huesped[5].Trim()), formatearnif(nif_acomp_Aci), formatearnif(huesped[3].Trim()), tipoturno, diasestancia, expedienteACi, noalojados, Convert.ToString(dr["age_nom_str"]), VectorExpedienteAnyo[1], HUE_GUID, huesped, nif_huesped, nif_original, Convert.ToString(dr["edi_des_str"]), huesped[14].Trim());
                }

                if (this.checkBox3.Checked)
                {

                    
                    try
                    {
                        this.misfuncionesBD.insert("AdvertenciaAsignadaReservas", "ADV_GUID,AAR_GUID,AAR_reg_bln,DEP_GUID,AAR_obs_str,AAR_alk_bln,AAR_con_bln", "'43','" + RES_GUID + "',0,0,' Plaza Adjudicada TV ',0,0");

                    }
                    catch
                    {


                    }



                }





            }

         



            string[] ExpedienteAno = huesped[0].Trim().Split(delimiters_re);
            this.misfuncionesBD.update("PresReservas", "res_DR_bon_str='" + ExpedienteAno[1].Trim() + "'", " (len(res_DR_bon_str)=0  OR res_DR_bon_str IS NULL OR res_DR_bon_str='') AND RES_GUID=" + RES_GUID);



            if (nif_original.Trim() == "")
            {

                RES_GUID = 0;

            }




            return RES_GUID;

        }






        //Existe prereserva para esa convocatoria

        public int ExisteReserva(int HUE_GUID, string[] huesped, string nif_huesped, string nif_original)
        {

            string nif = huesped[1].Trim();
            string rango_convocatorias = "";
            string estado = "";
            string numerodeadultos = "";
            string expedienteACi = "";
            DataSet miDataset = new DataSet();
            /*Añadimos ceros en el caso de que el nif sea menor de nueve digitos*/

            nif = formatearnif(nif);



            //Obtengo el rango de fechas de la convocatoria
            char[] delimiters = new char[] { '-' };
            rango_convocatorias = ObtenerRangoConvocatoria(Convert.ToDateTime(huesped[7].Trim()));
            string[] VectotConvocatorias = rango_convocatorias.Split(delimiters);


            char[] delimiters_expediente = new char[] { '/' };
            string[] VectorExpedienteAnyo = Convert.ToString(huesped[0].Trim()).Split(delimiters_expediente);


            int RES_GUID = 0;
            DataTable tabla_existe_registro = new DataTable();

            DateTime start = Convert.ToDateTime(VectotConvocatorias[0]);
            DateTime end = Convert.ToDateTime(VectotConvocatorias[0]);

            // lectura_fichero += System.Environment.NewLine + "SELECT RES_GUID FROM [PresReservas] WHERE HUE_GUID='" + Convert.ToString(HUE_GUID) + "' AND res_ent_dat between  convert(datetime, '" + Convert.ToDateTime(VectotConvocatorias[0]).ToString("dd/MM/yyyy") + "') AND  convert(datetime, '" + Convert.ToDateTime(VectotConvocatorias[1]).ToString("dd/MM/yyyy") + "')" + System.Environment.NewLine; 

            /*Obtenemos el id del ultimo contador*/
            string sql = "";


            if (this.checkBox3.Checked)
            {

           



                if (huesped[14] == "A")
                {

                    sql = "SELECT edi_des_str,age_nom_str,res_nal_bln,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[Reservas].RES_GUID,HuespedK.HUE_GUID,ReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,CASE res_est_int WHEN 0 THEN 'Reserva' WHEN 2 THEN 'Checkin/Checkout' WHEN 3 THEN 'No Show' WHEN 4 THEN 'Anulada' ELSE '' END AS Estado FROM [Reservas] INNER JOIN HuespedK ON HuespedK.HUE_GUID = Reservas.HUE_GUID LEFT  JOIN ReservaAcompanantes ON  ReservaAcompanantes.RES_GUID = Reservas.RES_GUID INNER JOIN Agencias ON  Agencias.AGE_GUID = Reservas.AGE_GUID  LEFT JOIN Edificios ON Edificios.EDI_GUID  = Reservas.EDI_GUID   WHERE  Reservas.AGE_GUID='" + this.agencia_balneario + "'   AND HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND  CTA_GUID IN ('" + System.Configuration.ConfigurationSettings.AppSettings["contrato_alta_tv"] + "') AND YEAR(res_ent_dat)=YEAR(GETDATE())  ";



                }
                else
                {


                    sql = "SELECT edi_des_str,age_nom_str,res_nal_bln,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[Reservas].RES_GUID,HuespedK.HUE_GUID,ReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,CASE res_est_int WHEN 0 THEN 'Reserva' WHEN 2 THEN 'Checkin/Checkout' WHEN 3 THEN 'No Show' WHEN 4 THEN 'Anulada' ELSE '' END AS Estado FROM [Reservas] INNER JOIN HuespedK ON HuespedK.HUE_GUID = Reservas.HUE_GUID LEFT  JOIN ReservaAcompanantes ON  ReservaAcompanantes.RES_GUID = Reservas.RES_GUID INNER JOIN Agencias ON  Agencias.AGE_GUID = Reservas.AGE_GUID  LEFT JOIN Edificios ON Edificios.EDI_GUID  = Reservas.EDI_GUID   WHERE  Reservas.AGE_GUID='" + this.agencia_balneario + "'   AND HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND  CTA_GUID IN ('" + System.Configuration.ConfigurationSettings.AppSettings["contrato_baja_tv"] + "') AND YEAR(res_ent_dat)=YEAR(GETDATE()) ";

                }





            }
            else
            {

                sql = "SELECT edi_des_str,age_nom_str,res_nal_bln,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[Reservas].RES_GUID,HuespedK.HUE_GUID,ReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,CASE res_est_int WHEN 0 THEN 'Reserva' WHEN 2 THEN 'Checkin/Checkout' WHEN 3 THEN 'No Show' WHEN 4 THEN 'Anulada' ELSE '' END AS Estado FROM [Reservas] INNER JOIN HuespedK ON HuespedK.HUE_GUID = Reservas.HUE_GUID LEFT  JOIN ReservaAcompanantes ON  ReservaAcompanantes.RES_GUID = Reservas.RES_GUID INNER JOIN Agencias ON  Agencias.AGE_GUID = Reservas.AGE_GUID  LEFT JOIN Edificios ON Edificios.EDI_GUID  = Reservas.EDI_GUID  WHERE  Reservas.AGE_GUID='" + this.agencia_balneario + "'  AND (  HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND res_ent_dat between '" + start.ToString(formato_datetime) + "' AND   '" + end.ToString(formato_datetime) + "') OR  (res_ent_dat between '" + Convert.ToDateTime("01/01/" + DateTime.Now.Year).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime("01/12/" + DateTime.Now.Year).ToString(formato_datetime) + "' AND  res_DR_bon_str='" + VectorExpedienteAnyo[1] + "' )  ";


            }




            //string sql_es = "SELECT res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[Reservas].RES_GUID,HuespedK.HUE_GUID,ReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,CASE res_est_int WHEN 0 THEN 'Reserva' WHEN 2 THEN 'Checkin/Checkout' WHEN 3 THEN 'No Show' WHEN 4 THEN 'Anulada' ELSE '' END AS Estado FROM [Reservas] INNER JOIN HuespedK ON HuespedK.HUE_GUID = Reservas.HUE_GUID LEFT  JOIN ReservaAcompanantes ON  ReservaAcompanantes.RES_GUID = Reservas.RES_GUID     WHERE (HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND res_ent_dat between '" + Convert.ToDateTime(VectotConvocatorias[0]).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime(VectotConvocatorias[1]).ToString(formato_datetime) + "')   OR  (res_ent_dat between '01/01/" + Convert.ToDateTime(VectotConvocatorias[1]).Year + "' AND   '31/12/" + Convert.ToDateTime(VectotConvocatorias[1]).Year + "' AND  res_DR_bon_str='" + VectorExpedienteAnyo[1] + "' ) ";

            miDataset = this.misfuncionesBD.obtenerDataSet(sql);




            tabla_existe_registro = miDataset.Tables[0];


            DateTime fechaReservaAci = DateTime.Now;
            string nif_huesped_aci = "";
            string IdAciHuesped = "";
            string idAcomp = "";
            string nif_acomp_Aci = "";
            string tipoturno = Convert.ToString(huesped[8].Trim());
            string diasestancia = "";
            string noalojados = "";
            char[] delimiters_re = new char[] { '/' };

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {
                RES_GUID = Convert.ToInt32(dr["RES_GUID"]);
                fechaReservaAci = Convert.ToDateTime(dr["res_ent_dat"]);
                nif_huesped_aci = formatearnif(Convert.ToString(dr["hue_nif_str"]));
                IdAciHuesped = Convert.ToString(dr["HUE_GUID"]);
                estado = Convert.ToString(dr["Estado"]);
                numerodeadultos = Convert.ToString(dr["res_spe_adu_int"]);
                idAcomp = Convert.ToString(dr["acomp"]);
                noalojados = Convert.ToString(dr["res_nal_bln"]);
                nif_acomp_Aci = ObtenerNifAcomAci(idAcomp);
                diasestancia = Convert.ToString(dr["DiferenciaDias"]);
                expedienteACi = Convert.ToString(dr["res_DR_bon_str"]);

                if (this.checkBox3.Checked)
                {

                    InsentarPacientesExcelAdvertenciasTV(Convert.ToDateTime(huesped[7].Trim()), fechaReservaAci, "Reserva", Convert.ToString(IdAciHuesped), Convert.ToString(huesped[1].Trim()), Convert.ToString(nif_huesped_aci), Convert.ToString(RES_GUID), VectorExpedienteAnyo[1], estado, numerodeadultos, Convert.ToString(huesped[5].Trim()), formatearnif(nif_acomp_Aci), formatearnif(huesped[3].Trim()), tipoturno, diasestancia, expedienteACi, noalojados, Convert.ToString(dr["age_nom_str"]), VectorExpedienteAnyo[1], HUE_GUID, huesped, nif_huesped, nif_original, Convert.ToString(dr["edi_des_str"]), huesped[14].Trim());

                }
                else
                {

                    InsentarPacientesExcelAdvertencias(Convert.ToDateTime(huesped[7].Trim()), fechaReservaAci, "Reserva", Convert.ToString(IdAciHuesped), Convert.ToString(huesped[1].Trim()), Convert.ToString(nif_huesped_aci), Convert.ToString(RES_GUID), VectorExpedienteAnyo[1], estado, numerodeadultos, Convert.ToString(huesped[5].Trim()), formatearnif(nif_acomp_Aci), formatearnif(huesped[3].Trim()), tipoturno, diasestancia, expedienteACi, noalojados, Convert.ToString(dr["age_nom_str"]), VectorExpedienteAnyo[1], HUE_GUID, huesped, nif_huesped, nif_original, Convert.ToString(dr["edi_des_str"]));
                }




                if (this.checkBox3.Checked)
                {

                    string[] ExpedienteAno = huesped[0].Trim().Split(delimiters_re);
                    this.misfuncionesBD.update("Reservas", "res_DR_bon_str='" + ExpedienteAno[1] + "'", "RES_GUID=" + RES_GUID);


                }


                if (this.checkBox3.Checked)
                {

                    //string[] ExpedienteAno = huesped[0].Trim().Split(delimiters_re);
                    try
                    {
                        this.misfuncionesBD.insert("AdvertenciaAsignadaReservas", "ADV_GUID,AAR_GUID,AAR_reg_bln,DEP_GUID,AAR_obs_str,AAR_alk_bln,AAR_con_bln", "'43','" + RES_GUID + "',0,0,' Plaza Adjudicada TV ',0,0");

                    }
                    catch
                    {


                    }


                }



            }



            string[] ExpedienteAno_R = huesped[0].Trim().Split(delimiters_re);
            this.misfuncionesBD.update("Reservas", "res_DR_bon_str='" + ExpedienteAno_R[1].Trim() + "'", " (len(res_DR_bon_str)=0  OR res_DR_bon_str IS NULL OR res_DR_bon_str='') AND RES_GUID=" + RES_GUID);


            if (nif_original.Trim() == "")
            {

                RES_GUID = 0;

            }




            return RES_GUID;

        }






        //obtenemos el rango de inicio y de fin de la convocatoria
        public string ObtenerRangoConvocatoria(DateTime fecha)
        {



            string rango = "";
            DateTime inicio = DateTime.Now;
            DateTime fin = DateTime.Now;

            DataTable tabla_existe_registro = new DataTable();
            DataSet miDataset = new DataSet();
            string sql = "";

            try
            {
                sql = "SELECT * FROM [convocatorias] where '" + fecha.ToString("MM/mm/yyyy") + "' between inicio and fin ";
                miDataset = this.misfuncionesBDExpedientes.obtenerDataSet(sql);
            }
            catch
            {

                sql = "SELECT * FROM [convocatorias] where '" + fecha.ToString("dd/MM/yyyy") + "' between inicio and fin ";
                miDataset = this.misfuncionesBDExpedientes.obtenerDataSet(sql);
            }

            tabla_existe_registro = miDataset.Tables[0];






            foreach (DataRow dr in tabla_existe_registro.Rows)
            {
                inicio = Convert.ToDateTime(dr["inicio"]);
                fin = Convert.ToDateTime(dr["fin"]);

            }



            rango = Convert.ToString(inicio).Replace("00:00:00", "").Replace("-", "/") + "-" + Convert.ToString(fin).Replace("00:00:00", "").Replace("-", "/");

            return rango;

        }


        //obtenemos el contrato asociado a esa fecha y turno
        public string ObtenerContrato(DateTime fecha, string tipo, string destino)
        {



            string contrato = "";
            string sqlcontrato = "";



            sqlcontrato = "SELECT * FROM contratos_imserso where fecha='" + Convert.ToDateTime(fecha).ToString("dd-MM-yyyy") + "' AND destino='" + Convert.ToString(destino) + "'";


            // MessageBox.Show(sqlcontrato);


            DataTable tabla_existe_registro = new DataTable();




            /*Obtenemos el id del ultimo contador*/
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable(sqlcontrato);




            foreach (DataRow dr in tabla_existe_registro.Rows)
            {




                if (tipo == "B (12 DIAS)")
                {

                    contrato = Convert.ToString(dr["oncedias"]);

                }


                if (tipo == "A (10 DIAS)")
                {
                    contrato = Convert.ToString(dr["nuevedias"]);

                }


                if (tipo == "A (7 DIAS)")
                {
                    contrato = Convert.ToString(dr["sietedias"]);

                }



            }



            //MessageBox.Show(contrato);

            return contrato;

        }



        public string formatearnif(string nif)
        {

            string formatearnif = "";
            int veces = 9 - nif.Length;


            // MessageBox.Show(Convert.ToString(nif.Length));

            /*Añadimos ceros en el caso de que el nif sea menor de nueve digitos*/
            if (nif.Length < 9)
            {



                for (int i = 1; i <= veces; i++)
                {

                    nif = "0" + nif;
                }

            }
            formatearnif = nif;

            return formatearnif;
        }

        private void button3_Click(object sender, EventArgs e)
        {

            try
            {
                this.ExtractData();
            }
            catch (Exception err)
            {
               
                SentrySdk.CaptureException(err);
               
                MessageBox.Show("Se ha producido un error");
                MessageBox.Show(err.ToString());
                this.Close();

            }




        }


        private void ExtractData()
        {

            this.id_balneario_seleccionado = Convert.ToString(this.comboBox1.SelectedValue);



            this.institucion = "Imserso";

            if (this.checkBox3.Checked == true)
            {

                this.institucion = "Generalitat";

            }

            if (Convert.ToString(comboBox1.SelectedItem) != "")
            {

                this.button1.Enabled = false;
                this.button2.Enabled = false;
                this.button3.Enabled = false;
                this.textBox1.Enabled = false;
                this.textBox3.Enabled = false;
                this.ResultBlock.Enabled = false;
                this.comboBox1.Enabled = false;

                lectura_fichero = "";
                lectura_fichero_nif = "";
                contador_lineas_fichero = 0;
                contador_lineas_fichero_dos = 0;

                DialogResult result = MessageBox.Show("¿Estas seguro del Balneario seleccionado?", "Salir", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)
                {
                    this.misfuncionesBD = new BDsqlserver.functionsBD(obtnener_balneario_conexion(Convert.ToString(this.comboBox1.SelectedValue)));
                    agencia_balneario = obtener_agencia(Convert.ToString(this.comboBox1.SelectedValue));
                    epr_GUID_tv = obtener_epr_GUID_tv(Convert.ToString(this.comboBox1.SelectedValue));
                    SPE_GUID = obtener_SPE_GUID(Convert.ToString(this.comboBox1.SelectedValue));
                    EDI_GUID = obtener_EDI_GUID(Convert.ToString(this.comboBox1.SelectedValue));
                    EDI_GUID_default = obtener_EDI_GUID_default(Convert.ToString(this.comboBox1.SelectedValue));
                    //MessageBox.Show(EDI_GUID);
                    PEN_GUID = this.obtener_pension(Convert.ToString(this.comboBox1.SelectedValue));
                    tipo_habitacion_balneario = this.obtener_tipo_habitacion(Convert.ToString(this.comboBox1.SelectedValue));
                    edificio_aci = this.obtener_tipo_edificio(Convert.ToString(this.comboBox1.SelectedValue));
                    formato_datetime = this.obtener_formatodatetime(Convert.ToString(this.comboBox1.SelectedValue));
                    backgroundWorker1.RunWorkerAsync();
                    this.contrato_balneario = valortexto_balneario(Convert.ToInt32(this.comboBox1.SelectedValue));
                    this.admin_balneario = obtener_balneario_admin(Convert.ToString(this.comboBox1.SelectedValue));
                    this.nacion = this.comboBox1.SelectedValue.ToString();
                    this.destino_FINAL = this.comboBox1.GetItemText(this.comboBox1.SelectedItem);
                    //MessageBox.Show(this.destino);


                    if (this.checkBox3.Checked)
                    {


                        this.ORR_GUID = obtener_ORR_GUID(Convert.ToString(this.comboBox1.SelectedValue), "tv");

                    }
                    else
                    {


                        this.ORR_GUID = obtener_ORR_GUID(Convert.ToString(this.comboBox1.SelectedValue), "imsesro");


                    }





                }

                else
                {

                    this.button1.Enabled = true;
                    this.button2.Enabled = true;
                    this.button3.Enabled = true;
                    this.textBox1.Enabled = true;
                    this.textBox3.Enabled = true;
                    this.ResultBlock.Enabled = true;
                    this.comboBox1.Enabled = true;
                }













            }
            else
            {

                MessageBox.Show("Debe seleccionar un balneario");

            }

        }


        private void button2_Click_1(object sender, EventArgs e)
        {


            int first = 0;
            textBox3.Clear();

            try
            {
                //System.IO.DriveInfo di = new System.IO.DriveInfo(textBox1.Text);
                //System.IO.DirectoryInfo dirInfo = DriveInfo.GetDrives();
                textBox3.Text = "";
                string directorio = textBox1.Text;
                //Console.WriteLine(dirInfo.Attributes.ToString());

                // Get the files in the directory and print out some information about them.
                string[] ficheros = Directory.GetFiles(directorio);

                foreach (String dir in ficheros)
                {
                    if (this.checkBox4.Checked)
                    {

                        if (dir.IndexOf(".xls") > 0 || dir.IndexOf(".xlsx") > 0)
                        {

                            textBox3.Text += dir + System.Environment.NewLine;

                        }

                    }
                    else
                    {


                        if (this.checkBox5.Checked)
                        {



                            if (dir.IndexOf(".csv") > 0 || dir.IndexOf(".csv") > 0)
                            {

                                textBox3.Text += dir + System.Environment.NewLine;

                            }



                        }
                        else
                        {



                            if (dir.IndexOf(".txt") > 0)
                            {

                                textBox3.Text += dir + System.Environment.NewLine;

                            }


                        }








                    }


                }
            }
            catch
            {


                MessageBox.Show("Debe seleccionar un directorio");

            }






        }


        //Insertamos en el Excel con las advertencias que correspondan



        public void InsentarPacientesExcelAdvertenciasTV(DateTime FechaImserso, DateTime FechaAci, string Tabla, string IdHuespedAci, string NifSolictenteImnserso, string NifHuespedAci, string IdRererva, string Expediente, string estado, string numeroadultos, string plazassolicitadas, string nif_acomp_Aci, string nif_acomp_imserso, string tipoturno, string diasestancia, string expedienteACi, string noalojados, string AgenciaBalneario, string res_bon_str, int HUE_GUID, string[] huesped, string nif_huesped, string nif_original, string edi_des_str, string tempora_tv)
        {
            string mesfichero = Convert.ToString(FechaImserso.Month);

            string anyofichero = Convert.ToString(FechaImserso.Year);
            string no_alojados_advertencias = "";

            string mes_entrada = Convert.ToString(FechaAci.Month);

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (this.checkBox3.Checked)
            {
                curFile = textBox1.Text + @"\AdvertenciasTV_" + Convert.ToString(FechaImserso.Year) + "\\Advertencias_" + fecha_actual_strg + ".xls";

            }
            else
            {


                curFile = textBox1.Text + @"\AdvertenciasImserso_" + mesfichero + "_" + anyofichero + "\\Advertencias_" + fecha_actual_strg + ".xls";
            }



            if (!File.Exists(curFile))
            {



                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "NifSol_ACI";
                xlWorkSheet.Cells[1, 2] = "NifSol_TV";
                xlWorkSheet.Cells[1, 3] = "NifAcom_ACI";
                xlWorkSheet.Cells[1, 4] = "NifAcom_TV";
                xlWorkSheet.Cells[1, 5] = "Expediente";
                xlWorkSheet.Cells[1, 6] = "Tipo";
                xlWorkSheet.Cells[1, 7] = "RESERVA";
                xlWorkSheet.Cells[1, 8] = "ADULTOS";
                //xlWorkSheet.Cells[1, 9] = "AGENCIA";
                xlWorkSheet.Cells[1, 9] = "ESTADO";
                xlWorkSheet.Cells[1, 10] = "ENTRADA_ACI";
                //xlWorkSheet.Cells[1, 12] = "ENTRADA_IMSERSO";
                //xlWorkSheet.Cells[1, 13] = "Solicitud";
                //xlWorkSheet.Cells[1, 11] = "DIAS_ESTANCIA";
                //xlWorkSheet.Cells[1, 15] = "TURNO_IMSERSO";
                xlWorkSheet.Cells[1, 11] = "Diferente_NIF_Titular";
                xlWorkSheet.Cells[1, 12] = "Plazas_Diferentes";
                xlWorkSheet.Cells[1, 13] = "Reserva_Anulada";
                xlWorkSheet.Cells[1, 14] = "Dias_Estancia_Diferentes";
                xlWorkSheet.Cells[1, 15] = "Diferente_NIF_Acompañante";
                //xlWorkSheet.Cells[1, 17] = "Fecha_Entrada_Diferente";
                //xlWorkSheet.Cells[1, 20] = "Prereserva_Estado_No";
                xlWorkSheet.Cells[1, 16] = "Comprobar_Nombre_Apellidos_Titular";
                xlWorkSheet.Cells[1, 17] = "Fecha_Nacimiento_Erronea_Titular";
                xlWorkSheet.Cells[1, 18] = "Comprobar_Nombre_Apellidos_Acompanate";
                xlWorkSheet.Cells[1, 19] = "Fecha_Nacimiento_Erronea_Acompanate";
                xlWorkSheet.Cells[1, 20] = "Temporada_Erronea";
                xlWorkSheet.Cells[1, 21] = "Balneario";
                xlWorkSheet.Cells[1, 22] = "Hotel";


                // xlWorkSheet.Cells[1, 17] = "OBSERVACION";

                xlWorkBook.SaveAs(curFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);


                //MessageBox.Show("No existe");
            }
            else
            {

                //MessageBox.Show("existe");


            }


            this.fichero_advertencias = "";

            string sql = "";





            CAMPOS_excel_fichero = "NifSol_ACI,NifSol_TV,NifAcom_ACI,NifAcom_TV,Expediente,Tipo,RESERVA,ADULTOS,ESTADO,ENTRADA_ACI,Balneario,Hotel,Diferente_NIF_Titular,Plazas_Diferentes,Reserva_Anulada,Dias_Estancia_Diferentes,Diferente_NIF_Acompañante,Comprobar_Nombre_Apellidos_Titular,Fecha_Nacimiento_Erronea_Titular,Comprobar_Nombre_Apellidos_Acompanate,Fecha_Nacimiento_Erronea_Acompanate,Temporada_Erronea";


            sql = "Insert into [Hoja1$] (" + CAMPOS_excel_fichero + ") values('" + NifHuespedAci.Trim() + "','" + formatearnif(NifSolictenteImnserso.Trim()) + "','" + nif_acomp_Aci.Trim() + "','" + formatearnif(nif_acomp_imserso.Trim()) + "','" + Expediente.Trim() + "','" + Tabla.Trim() + "','" + IdRererva.Trim() + "','" + numeroadultos.Trim() + "','" + estado.Trim() + "','" + FechaAci.ToString("dd/MM/yyyy") + "','" + this.destino_FINAL + "','" + edi_des_str + "'";


            bool advertencia_ok = false;


            string advertencia = "NO";
            if (formatearnif(NifSolictenteImnserso.Trim()) != formatearnif(NifHuespedAci.Trim()))
            {
                if (plazassolicitadas != "B")
                {
                    advertencia = "SI";
                    advertencia_ok = true;
                }
                else
                {


                    if (formatearnif(nif_acomp_imserso.Trim()) != formatearnif(NifHuespedAci.Trim()))
                    {


                        advertencia = "SI";
                        advertencia_ok = true;

                    }


                }
            }
            sql += ",'" + advertencia + "'";






            string nif = huesped[1].Trim();
            string rango_convocatorias = "";




            nif = formatearnif(nif);
            char[] delimiters = new char[] { '-' };
            rango_convocatorias = ObtenerRangoConvocatoria(Convert.ToDateTime(huesped[7].Trim()));
            string[] VectotConvocatorias = rango_convocatorias.Split(delimiters);

            char[] delimiters_expediente = new char[] { '/' };
            string[] VectorExpedienteAnyo = Convert.ToString(huesped[0].Trim()).Split(delimiters_expediente);


            DataTable tabla_existe_registro = new DataTable();
            DataSet miDataset = new DataSet();

            string sql_reserva = "";




            advertencia = "NO";
            if (estado == "Anulada")
            {


                //Tenemos que comprobar que no tenga una reserva en estado RESERVA O  CHECK


                if (this.checkBox3.Checked)
                {




                    sql_reserva = "SELECT age_nom_str,res_nal_bln,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[Reservas].RES_GUID,HuespedK.HUE_GUID,ReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,CASE res_est_int WHEN 0 THEN 'Reserva' WHEN 2 THEN 'Checkin/Checkout' WHEN 3 THEN 'No Show' WHEN 4 THEN 'Anulada' ELSE '' END AS Estado FROM [Reservas] INNER JOIN HuespedK ON HuespedK.HUE_GUID = Reservas.HUE_GUID LEFT  JOIN ReservaAcompanantes ON  ReservaAcompanantes.RES_GUID = Reservas.RES_GUID INNER JOIN Agencias ON  Agencias.AGE_GUID = Reservas.AGE_GUID    WHERE  Reservas.AGE_GUID='" + this.agencia_balneario + "'  AND res_est_int  IN (0,2) AND HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND  CTA_GUID IN ('" + System.Configuration.ConfigurationSettings.AppSettings["contrato_baja_tv"] + "') ";





                }
                else
                {

                    sql_reserva = "SELECT age_nom_str,res_nal_bln,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[Reservas].RES_GUID,HuespedK.HUE_GUID,ReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,CASE res_est_int WHEN 0 THEN 'Reserva' WHEN 2 THEN 'Checkin/Checkout' WHEN 3 THEN 'No Show' WHEN 4 THEN 'Anulada' ELSE '' END AS Estado FROM [Reservas] INNER JOIN HuespedK ON HuespedK.HUE_GUID = Reservas.HUE_GUID LEFT  JOIN ReservaAcompanantes ON  ReservaAcompanantes.RES_GUID = Reservas.RES_GUID INNER JOIN Agencias ON  Agencias.AGE_GUID = Reservas.AGE_GUID  WHERE  Reservas.AGE_GUID='" + this.agencia_balneario + "'   AND  res_est_int IN (0,2) AND HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND res_ent_dat between '" + Convert.ToDateTime(VectotConvocatorias[0]).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime(VectotConvocatorias[1]).ToString(formato_datetime) + "' ";


                }


                miDataset = this.misfuncionesBD.obtenerDataSet(sql_reserva);
                tabla_existe_registro = miDataset.Tables[0];
                bool reserva_ok = false;
                foreach (DataRow dr in tabla_existe_registro.Rows)
                {

                    reserva_ok = true;
                    DateTime fecha_reserva_ok = Convert.ToDateTime(dr["res_ent_dat"]);
                    diasestancia = Convert.ToString(dr["DiferenciaDias"]);
                    FechaAci = Convert.ToDateTime(dr["res_ent_dat"]);

                }

                if (!reserva_ok)
                {
                    advertencia = "SI";
                    advertencia_ok = true;


                }




            }
            sql += ",'" + advertencia + "'";


            advertencia = "NO";
            if ((diasestancia.Trim() == "11" && tipoturno.Trim() != "B (12 DIAS)") || ((diasestancia.Trim() == "9" && tipoturno.Trim() != "A (10 DIAS)")) || (diasestancia.Trim() == "7" && tipoturno.Trim() != "A (7 DIAS)"))
            {
                advertencia = "SI";
                advertencia_ok = true;

            }
            sql += ",'" + advertencia + "'";





            advertencia = "NO";
            if (plazassolicitadas == "C")
            {
                if (formatearnif(nif_acomp_Aci.Trim()) != formatearnif(nif_acomp_imserso.Trim()))
                {
                    advertencia = "SI";
                    advertencia_ok = true;
                }

            }
            sql += ",'" + advertencia + "'";









            tabla_existe_registro = new DataTable();


            string sql_prererva = "";

            // MessageBox.Show(estado.Trim());

            if (this.checkBox3.Checked)
            {

                //sql = "SELECT age_nom_str,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[PresReservas].RES_GUID,HuespedK.HUE_GUID,PresReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,PreReservaEstados.epr_des_str as Estado,prr_est_byt FROM PresReservas INNER JOIN  PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID INNER JOIN HuespedK ON HuespedK.HUE_GUID = PresReservas.HUE_GUID LEFT  JOIN  PresReservaAcompanantes ON  PresReservaAcompanantes.RES_GUID = PresReservas.RES_GUID   INNER JOIN Agencias ON  Agencias.AGE_GUID = PresReservas.AGE_GUID  WHERE  PresReservas.AGE_GUID='" + this.agencia_balneario + "' AND HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND (res_ent_dat between '" + Convert.ToDateTime("01/01/" + Convert.ToDateTime(VectotConvocatorias[1]).Year).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime("31/12/" + Convert.ToDateTime(VectotConvocatorias[1]).Year).ToString(formato_datetime) + "' AND  res_DR_bon_str='" + VectorExpedienteAnyo[1] + "' ) ";
                sql_prererva = "SELECT age_nom_str,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[PresReservas].RES_GUID,HuespedK.HUE_GUID,PresReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,PreReservaEstados.epr_des_str as Estado,prr_est_byt FROM PresReservas INNER JOIN  PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID INNER JOIN HuespedK ON HuespedK.HUE_GUID = PresReservas.HUE_GUID LEFT  JOIN  PresReservaAcompanantes ON  PresReservaAcompanantes.RES_GUID = PresReservas.RES_GUID   INNER JOIN Agencias ON  Agencias.AGE_GUID = PresReservas.AGE_GUID  WHERE  PresReservas.AGE_GUID='" + this.agencia_balneario + "' AND HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND (res_ent_dat between '" + Convert.ToDateTime("01/01/" + Convert.ToDateTime(VectotConvocatorias[1]).Year).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime("31/12/" + Convert.ToDateTime(VectotConvocatorias[1]).Year).ToString(formato_datetime) + "'  ) ";

            }
            else
            {

                sql_prererva = "SELECT age_nom_str,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[PresReservas].RES_GUID,HuespedK.HUE_GUID,PresReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,PreReservaEstados.epr_des_str as Estado,prr_est_byt FROM PresReservas INNER JOIN  PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID INNER JOIN HuespedK ON HuespedK.HUE_GUID = PresReservas.HUE_GUID LEFT  JOIN  PresReservaAcompanantes ON  PresReservaAcompanantes.RES_GUID = PresReservas.RES_GUID   INNER JOIN Agencias ON  Agencias.AGE_GUID = PresReservas.AGE_GUID  WHERE PresReservas.AGE_GUID='" + this.agencia_balneario + "' AND (HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND res_ent_dat between '" + Convert.ToDateTime(VectotConvocatorias[0]).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime(VectotConvocatorias[1]).ToString(formato_datetime) + "')  OR (res_ent_dat between '" + Convert.ToDateTime("01/01/" + DateTime.Now.Year).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime("31/12/" + DateTime.Now.Year).ToString(formato_datetime) + "' AND  res_DR_bon_str='" + VectorExpedienteAnyo[1] + "' ) ";


            }

            miDataset = this.misfuncionesBD.obtenerDataSet(sql_prererva);
            tabla_existe_registro = miDataset.Tables[0];
            bool prereserva_ok = false;
            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                prereserva_ok = true;
                estado = Convert.ToString(dr["Estado"]);

            }





            advertencia = "NO";

            if (estado.ToUpper() == "NO")
            {

                advertencia = "SI";
                advertencia_ok = true;

            }
            sql += ",'" + advertencia + "'";





            // MessageBox.Show(estado.Trim());



            //Comprobamos si nombre titular es compuesto
            char[] delimiters_nombre = new char[] { ' ' };
            string[] array_nombre = Convert.ToString(huesped[2].Trim()).Split(delimiters_nombre);
            advertencia = "NO";


            if (array_nombre.Length > 3)
            {

                advertencia = "SI";
                advertencia_ok = true;

            }
            sql += ",'" + advertencia + "'";



            // MessageBox.Show(huesped[30]);


            //Comprobamos si la fecha nacimiento titular
            advertencia = "NO";

            if (!CheckBird(Convert.ToString(nif_huesped)))
            {

                advertencia = "SI";
                advertencia_ok = true;
            }



            sql += ",'" + advertencia + "'";



            //Comprobamos si nombre acompañente

            advertencia = "NO";
            array_nombre = Convert.ToString(huesped[4].Trim()).Split(delimiters_nombre);


            if (plazassolicitadas == "C")
            {
                if (array_nombre.Length > 3)
                {

                    advertencia = "SI";
                    advertencia_ok = true;

                }
            }

            sql += ",'" + advertencia + "'";



            advertencia = "NO";

            //Comprobamos si la fecha nacimiento acompañente
            if (plazassolicitadas == "C")
            {
                /*try
                {
                    Convert.ToDateTime(huesped[31]);
                    
                }
                catch
                {

                    advertencia = "SI";
                    advertencia_ok = true;
                }*/


                if (!CheckBird(Convert.ToString(nif_acomp_imserso)))
                {

                    advertencia = "SI";
                    advertencia_ok = true;
                }




            }




            sql += ",'" + advertencia + "'";



            //Temporada 

            if (this.checkBox3.Checked)
            {

                advertencia = "NO";

                string[] temporada_baja = { "1", "2", "3", "4", "5", "6", "7", "11", "12" };
                string[] temporada_alta = { "8", "9", "10" };

                if (tempora_tv == "B" && temporada_alta.Contains(mes_entrada))
                {

                    advertencia = "SI";

                }


                if (tempora_tv == "A" && temporada_baja.Contains(mes_entrada))
                {

                    advertencia = "SI";

                }

                /* if (estado.ToUpper() == "NO")
                 {

                     advertencia = "SI";
                     advertencia_ok = true;

                 }*/
                sql += ",'" + advertencia + "'";



            }
            else
            {

                advertencia = "NO";

            }










            sql += ")";



            if (advertencia_ok)
            {

                // MessageBox.Show(advertencia_ok.ToString());

                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();

                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + curFile + ";Extended Properties=Excel 12.0;");
                MyConnection.Open();
                myCommand.Connection = MyConnection;

                myCommand.CommandText = sql;
                myCommand.ExecuteNonQuery();
                MyConnection.Close();
            }






        }

        public void InsentarPacientesExcelAdvertencias(DateTime FechaImserso, DateTime FechaAci, string Tabla, string IdHuespedAci, string NifSolictenteImnserso, string NifHuespedAci, string IdRererva, string Expediente, string estado, string numeroadultos, string plazassolicitadas, string nif_acomp_Aci, string nif_acomp_imserso, string tipoturno, string diasestancia, string expedienteACi, string noalojados, string AgenciaBalneario, string res_bon_str, int HUE_GUID, string[] huesped, string nif_huesped, string nif_original, string edi_des_str)
        {
            string mesfichero = Convert.ToString(FechaImserso.Month);
            string anyofichero = Convert.ToString(FechaImserso.Year);
            string no_alojados_advertencias = "";

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (this.checkBox3.Checked)
            {
                curFile = textBox1.Text + @"\AdvertenciasTV_" + Convert.ToString(FechaImserso.Year) + "\\Advertencias_" + fecha_actual_strg + ".xls";

            }
            else
            {


                curFile = textBox1.Text + @"\AdvertenciasImserso_" + mesfichero + "_" + anyofichero + "\\Advertencias_" + fecha_actual_strg + ".xls";
            }



            if (!File.Exists(curFile))
            {



                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "NifSol_ACI";
                xlWorkSheet.Cells[1, 2] = "NifSol_IMSERSO";
                xlWorkSheet.Cells[1, 3] = "NifAcom_ACI";
                xlWorkSheet.Cells[1, 4] = "NifAcom_IMSERSO";
                xlWorkSheet.Cells[1, 5] = "Expediente";
                xlWorkSheet.Cells[1, 6] = "Tipo";
                xlWorkSheet.Cells[1, 7] = "RESERVA";
                xlWorkSheet.Cells[1, 8] = "ADULTOS";
                xlWorkSheet.Cells[1, 9] = "AGENCIA";
                xlWorkSheet.Cells[1, 10] = "ESTADO";
                xlWorkSheet.Cells[1, 11] = "ENTRADA_ACI";
                xlWorkSheet.Cells[1, 12] = "ENTRADA_IMSERSO";
                xlWorkSheet.Cells[1, 13] = "Solicitud";
                xlWorkSheet.Cells[1, 14] = "DIAS_ESTANCIA";
                xlWorkSheet.Cells[1, 15] = "TURNO_IMSERSO";
                xlWorkSheet.Cells[1, 16] = "Diferente_NIF_Titular";
                xlWorkSheet.Cells[1, 17] = "Plazas_Diferentes";
                xlWorkSheet.Cells[1, 18] = "Reserva_Anulada";
                xlWorkSheet.Cells[1, 19] = "Dias_Estancia_Diferentes";
                xlWorkSheet.Cells[1, 20] = "Diferente_NIF_Acompañante";
                xlWorkSheet.Cells[1, 21] = "Fecha_Entrada_Diferente";
                xlWorkSheet.Cells[1, 22] = "Prereserva_Estado_No";
                xlWorkSheet.Cells[1, 23] = "Comprobar_Nombre_Apellidos_Titular";
                xlWorkSheet.Cells[1, 24] = "Fecha_Nacimiento_Erronea_Titular";
                xlWorkSheet.Cells[1, 25] = "Comprobar_Nombre_Apellidos_Acompanate";
                xlWorkSheet.Cells[1, 26] = "Fecha_Nacimiento_Erronea_Acompanate";
                xlWorkSheet.Cells[1, 27] = "Balneario";
                xlWorkSheet.Cells[1, 28] = "Hotel";


                // xlWorkSheet.Cells[1, 17] = "OBSERVACION";

                //MessageBox.Show(curFile);
                xlWorkBook.SaveAs(curFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);


                //MessageBox.Show("No existe");
            }
            else
            {

                //MessageBox.Show("existe");


            }


            this.fichero_advertencias = "";

            string sql = "";





            CAMPOS_excel_fichero = "NifSol_ACI,NifSol_IMSERSO,NifAcom_ACI,NifAcom_IMSERSO,Expediente,Tipo,RESERVA,ADULTOS,AGENCIA,ESTADO,ENTRADA_ACI,ENTRADA_IMSERSO,Solicitud,DIAS_ESTANCIA,TURNO_IMSERSO,Balneario,Hotel,Diferente_NIF_Titular,Plazas_Diferentes,Reserva_Anulada,Dias_Estancia_Diferentes,Diferente_NIF_Acompañante,Fecha_Entrada_Diferente,Prereserva_Estado_No,Comprobar_Nombre_Apellidos_Titular,Fecha_Nacimiento_Erronea_Titular,Comprobar_Nombre_Apellidos_Acompanate,Fecha_Nacimiento_Erronea_Acompanate";


            sql = "Insert into [Hoja1$] (" + CAMPOS_excel_fichero + ") values('" + NifHuespedAci.Trim() + "','" + formatearnif(NifSolictenteImnserso.Trim()) + "','" + nif_acomp_Aci.Trim() + "','" + formatearnif(nif_acomp_imserso.Trim()) + "','" + Expediente.Trim() + "','" + Tabla.Trim() + "','" + IdRererva.Trim() + "','" + numeroadultos.Trim() + "','" + AgenciaBalneario.Trim() + "','" + estado.Trim() + "','" + FechaAci.ToString("dd/MM/yyyy") + "','" + FechaImserso.ToString("dd/MM/yyyy") + "','" + plazassolicitadas + "'," + diasestancia.Trim() + ",'" + tipoturno + "','" + this.destino_FINAL + "','" + edi_des_str + "'";


            bool advertencia_ok = false;


            string advertencia = "NO";
            if (formatearnif(NifSolictenteImnserso.Trim()) != formatearnif(NifHuespedAci.Trim()))
            {
                if (plazassolicitadas != "B")
                {
                    advertencia = "SI";
                    advertencia_ok = true;
                }
                else
                {


                    if (formatearnif(nif_acomp_imserso.Trim()) != formatearnif(NifHuespedAci.Trim()))
                    {


                        advertencia = "SI";
                        advertencia_ok = true;

                    }


                }
            }
            sql += ",'" + advertencia + "'";


            advertencia = "NO";
            if ((numeroadultos == "1" && plazassolicitadas == "C") || (numeroadultos == "2" && plazassolicitadas == "B") || (numeroadultos == "2" && plazassolicitadas == "A"))

            {
                advertencia = "SI";
                advertencia_ok = true;

            }
            sql += ",'" + advertencia + "'";




            ///COmprobamos si tiene la prereserva e estado NO



            string nif = huesped[1].Trim();
            string rango_convocatorias = "";

            //string expedienteACi = "";



            nif = formatearnif(nif);
            char[] delimiters = new char[] { '-' };
            rango_convocatorias = ObtenerRangoConvocatoria(Convert.ToDateTime(huesped[7].Trim()));
            string[] VectotConvocatorias = rango_convocatorias.Split(delimiters);

            char[] delimiters_expediente = new char[] { '/' };
            string[] VectorExpedienteAnyo = Convert.ToString(huesped[0].Trim()).Split(delimiters_expediente);


            DataTable tabla_existe_registro = new DataTable();
            DataSet miDataset = new DataSet();

            string sql_reserva = "";




            advertencia = "NO";
            if (estado == "Anulada")
            {


                //Tenemos que comprobar que no tenga una reserva en estado RESERVA O  CHECK


                if (this.checkBox3.Checked)
                {




                    sql_reserva = "SELECT age_nom_str,res_nal_bln,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[Reservas].RES_GUID,HuespedK.HUE_GUID,ReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,CASE res_est_int WHEN 0 THEN 'Reserva' WHEN 2 THEN 'Checkin/Checkout' WHEN 3 THEN 'No Show' WHEN 4 THEN 'Anulada' ELSE '' END AS Estado FROM [Reservas] INNER JOIN HuespedK ON HuespedK.HUE_GUID = Reservas.HUE_GUID LEFT  JOIN ReservaAcompanantes ON  ReservaAcompanantes.RES_GUID = Reservas.RES_GUID INNER JOIN Agencias ON  Agencias.AGE_GUID = Reservas.AGE_GUID    WHERE  Reservas.AGE_GUID='" + this.agencia_balneario + "'  AND res_est_int  IN (0,2) AND HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND  CTA_GUID IN ('" + System.Configuration.ConfigurationSettings.AppSettings["contrato_baja_tv"] + "') ";





                }
                else
                {

                    sql_reserva = "SELECT age_nom_str,res_nal_bln,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[Reservas].RES_GUID,HuespedK.HUE_GUID,ReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,CASE res_est_int WHEN 0 THEN 'Reserva' WHEN 2 THEN 'Checkin/Checkout' WHEN 3 THEN 'No Show' WHEN 4 THEN 'Anulada' ELSE '' END AS Estado FROM [Reservas] INNER JOIN HuespedK ON HuespedK.HUE_GUID = Reservas.HUE_GUID LEFT  JOIN ReservaAcompanantes ON  ReservaAcompanantes.RES_GUID = Reservas.RES_GUID INNER JOIN Agencias ON  Agencias.AGE_GUID = Reservas.AGE_GUID  WHERE  Reservas.AGE_GUID='" + this.agencia_balneario + "'   AND  res_est_int IN (0,2) AND HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND res_ent_dat between '" + Convert.ToDateTime(VectotConvocatorias[0]).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime(VectotConvocatorias[1]).ToString(formato_datetime) + "' ";


                }


                miDataset = this.misfuncionesBD.obtenerDataSet(sql_reserva);
                tabla_existe_registro = miDataset.Tables[0];
                bool reserva_ok = false;
                foreach (DataRow dr in tabla_existe_registro.Rows)
                {

                    reserva_ok = true;
                    DateTime fecha_reserva_ok = Convert.ToDateTime(dr["res_ent_dat"]);
                    diasestancia = Convert.ToString(dr["DiferenciaDias"]);
                    FechaAci = Convert.ToDateTime(dr["res_ent_dat"]);

                }

                if (!reserva_ok)
                {
                    advertencia = "SI";
                    advertencia_ok = true;


                }




            }
            sql += ",'" + advertencia + "'";


            advertencia = "NO";
            if ((diasestancia.Trim() == "11" && tipoturno.Trim() != "B (12 DIAS)") || ((diasestancia.Trim() == "9" && tipoturno.Trim() != "A (10 DIAS)")) || (diasestancia.Trim() == "7" && tipoturno.Trim() != "A (7 DIAS)"))
            {
                advertencia = "SI";
                advertencia_ok = true;

            }
            sql += ",'" + advertencia + "'";





            advertencia = "NO";
            if (plazassolicitadas == "C")
            {
                if (formatearnif(nif_acomp_Aci.Trim()) != formatearnif(nif_acomp_imserso.Trim()))
                {
                    advertencia = "SI";
                    advertencia_ok = true;
                }

            }
            sql += ",'" + advertencia + "'";



            advertencia = "NO";
            if (FechaImserso.CompareTo(FechaAci) != 0)
            {

                advertencia = "SI";
                advertencia_ok = true;


            }
            sql += ",'" + advertencia + "'";



            ///COmprobamos si tiene la prereserva e estado NO






            tabla_existe_registro = new DataTable();


            string sql_prererva = "";

            // MessageBox.Show(estado.Trim());

            if (this.checkBox3.Checked)
            {

                //sql = "SELECT age_nom_str,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[PresReservas].RES_GUID,HuespedK.HUE_GUID,PresReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,PreReservaEstados.epr_des_str as Estado,prr_est_byt FROM PresReservas INNER JOIN  PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID INNER JOIN HuespedK ON HuespedK.HUE_GUID = PresReservas.HUE_GUID LEFT  JOIN  PresReservaAcompanantes ON  PresReservaAcompanantes.RES_GUID = PresReservas.RES_GUID   INNER JOIN Agencias ON  Agencias.AGE_GUID = PresReservas.AGE_GUID  WHERE  PresReservas.AGE_GUID='" + this.agencia_balneario + "' AND HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND (res_ent_dat between '" + Convert.ToDateTime("01/01/" + Convert.ToDateTime(VectotConvocatorias[1]).Year).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime("31/12/" + Convert.ToDateTime(VectotConvocatorias[1]).Year).ToString(formato_datetime) + "' AND  res_DR_bon_str='" + VectorExpedienteAnyo[1] + "' ) ";
                sql_prererva = "SELECT age_nom_str,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[PresReservas].RES_GUID,HuespedK.HUE_GUID,PresReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,PreReservaEstados.epr_des_str as Estado,prr_est_byt FROM PresReservas INNER JOIN  PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID INNER JOIN HuespedK ON HuespedK.HUE_GUID = PresReservas.HUE_GUID LEFT  JOIN  PresReservaAcompanantes ON  PresReservaAcompanantes.RES_GUID = PresReservas.RES_GUID   INNER JOIN Agencias ON  Agencias.AGE_GUID = PresReservas.AGE_GUID  WHERE  PresReservas.AGE_GUID='" + this.agencia_balneario + "' AND HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND (res_ent_dat between '" + Convert.ToDateTime("01/01/" + Convert.ToDateTime(VectotConvocatorias[1]).Year).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime("31/12/" + Convert.ToDateTime(VectotConvocatorias[1]).Year).ToString(formato_datetime) + "'  ) ";

            }
            else
            {

                sql_prererva = "SELECT age_nom_str,res_DR_bon_str,DateDiff(day, res_ent_dat, res_sal_dat) as DiferenciaDias,res_spe_adu_int,[PresReservas].RES_GUID,HuespedK.HUE_GUID,PresReservaAcompanantes.HUE_GUID as acomp,hue_nif_str,res_ent_dat,PreReservaEstados.epr_des_str as Estado,prr_est_byt FROM PresReservas INNER JOIN  PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID INNER JOIN HuespedK ON HuespedK.HUE_GUID = PresReservas.HUE_GUID LEFT  JOIN  PresReservaAcompanantes ON  PresReservaAcompanantes.RES_GUID = PresReservas.RES_GUID   INNER JOIN Agencias ON  Agencias.AGE_GUID = PresReservas.AGE_GUID  WHERE PresReservas.AGE_GUID='" + this.agencia_balneario + "' AND (HuespedK.hue_nif_str='" + Convert.ToString(nif_huesped) + "' AND res_ent_dat between '" + Convert.ToDateTime(VectotConvocatorias[0]).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime(VectotConvocatorias[1]).ToString(formato_datetime) + "')  OR (res_ent_dat between '" + Convert.ToDateTime("01/01/" + DateTime.Now.Year).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime("31/12/" + DateTime.Now.Year).ToString(formato_datetime) + "' AND  res_DR_bon_str='" + VectorExpedienteAnyo[1] + "' ) ";


            }

            miDataset = this.misfuncionesBD.obtenerDataSet(sql_prererva);
            tabla_existe_registro = miDataset.Tables[0];
            bool prereserva_ok = false;
            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                prereserva_ok = true;
                estado = Convert.ToString(dr["Estado"]);

            }



            /* advertencia = "NO";

             if (!prereserva_ok)
             {

                 advertencia = "SI";
                 advertencia_ok = true;

             }
             sql += ",'" + advertencia + "'";*/



            advertencia = "NO";

            if (estado.ToUpper() == "NO")
            {

                advertencia = "SI";
                advertencia_ok = true;

            }
            sql += ",'" + advertencia + "'";

            // MessageBox.Show(estado.Trim());



            //Comprobamos si nombre titular es compuesto
            char[] delimiters_nombre = new char[] { ' ' };
            string[] array_nombre = Convert.ToString(huesped[2].Trim()).Split(delimiters_nombre);
            advertencia = "NO";


            if (array_nombre.Length > 3)
            {

                advertencia = "SI";
                advertencia_ok = true;

            }
            sql += ",'" + advertencia + "'";



            // MessageBox.Show(huesped[30]);


            //Comprobamos si la fecha nacimiento titular
            advertencia = "NO";

            if (!CheckBird(Convert.ToString(nif_huesped)))
            {

                advertencia = "SI";
                advertencia_ok = true;
            }



            sql += ",'" + advertencia + "'";



            //Comprobamos si nombre acompañente

            advertencia = "NO";
            array_nombre = Convert.ToString(huesped[4].Trim()).Split(delimiters_nombre);


            if (plazassolicitadas == "C")
            {
                if (array_nombre.Length > 3)
                {

                    advertencia = "SI";
                    advertencia_ok = true;

                }
            }

            sql += ",'" + advertencia + "'";



            advertencia = "NO";

            //Comprobamos si la fecha nacimiento acompañente
            if (plazassolicitadas == "C")
            {
                /*try
                {
                    Convert.ToDateTime(huesped[31]);
                    
                }
                catch
                {

                    advertencia = "SI";
                    advertencia_ok = true;
                }*/


                if (!CheckBird(Convert.ToString(nif_acomp_imserso)))
                {

                    advertencia = "SI";
                    advertencia_ok = true;
                }




            }




            sql += ",'" + advertencia + "'";







            sql += ")";



            if (advertencia_ok)
            {

                // MessageBox.Show(advertencia_ok.ToString());

                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();

                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + curFile + ";Extended Properties=Excel 12.0;");
                MyConnection.Open();
                myCommand.Connection = MyConnection;

                myCommand.CommandText = sql;
                myCommand.ExecuteNonQuery();
                MyConnection.Close();
            }






        }




        //obtener id ACI acompañante
        public string ObtenerNifAcomAci(string idaci)
        {


            string nif = "";

            DataTable tabla_existe_registro = new DataTable();

            /*Obtenemos el id del ultimo contador*/
            tabla_existe_registro = this.misfuncionesBD.obtenerDatable("SELECT hue_nif_str FROM [HuespedK] where HUE_GUID='" + idaci + "'");




            foreach (DataRow dr in tabla_existe_registro.Rows)
            {
                nif = Convert.ToString(dr["hue_nif_str"]);

            }

            return nif;

        }



        //Saca todos los expedientes que no estan en Imserso y los inserta en el Excel
        public void HuespedesNoEstanTVExcel(string mes, string year_fichero, ArrayList ListaNifsImsersoMes)
        {


            string nif = "";
            string contratos_imsesro = "";
            ArrayList expediente_estado = new ArrayList();
            string no_alojados_advertencias = "";

            string mesfichero = Convert.ToString(mes);
            string anyofichero = Convert.ToString(year_fichero);



            CAMPOS_excel_fichero = "NifSol_ACI,NifSol_TV,NifAcom_ACI,NifAcom_TV,Expediente,Tipo,RESERVA,ADULTOS,ESTADO,ENTRADA_ACI,Balneario,Hotel,Diferente_NIF_Titular,Plazas_Diferentes,Reserva_Anulada,Dias_Estancia_Diferentes,Diferente_NIF_Acompañante,Comprobar_Nombre_Apellidos_Titular,Fecha_Nacimiento_Erronea_Titular,Comprobar_Nombre_Apellidos_Acompanate,Fecha_Nacimiento_Erronea_Acompanate";


            //sql = "Insert into [Hoja1$] (" + CAMPOS_excel_fichero + ") values('" + NifHuespedAci.Trim() + "','" + formatearnif(NifSolictenteImnserso.Trim()) + "','" + nif_acomp_Aci.Trim() + "','" + formatearnif(nif_acomp_imserso.Trim()) + "','" + Expediente.Trim() + "','" + Tabla.Trim() + "','" + IdRererva.Trim() + "','" + numeroadultos.Trim() + "','" + estado.Trim() + "','" + FechaAci.ToString("dd/MM/yyyy") + "','" + this.destino_FINAL + "','" + edi_des_str + "'";


            //Obtenemos los contratos

            DataTable tabla_existe_registro_contratos = new DataTable();

            /*Obtenemos el id del ultimo contador*/




            if (!File.Exists(curFile))
            {


                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();



                curFile = textBox1.Text + @"\AdvertenciasTV_" + Convert.ToString(anyofichero) + "\\Advertencias_" + fecha_actual_strg + ".xls";



                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "NifSol_ACI";
                xlWorkSheet.Cells[1, 2] = "NifSol_TV";
                xlWorkSheet.Cells[1, 3] = "NifAcom_ACI";
                xlWorkSheet.Cells[1, 4] = "NifAcom_TV";
                xlWorkSheet.Cells[1, 5] = "Expediente";
                xlWorkSheet.Cells[1, 6] = "Tipo";
                xlWorkSheet.Cells[1, 7] = "RESERVA";
                xlWorkSheet.Cells[1, 8] = "ADULTOS";
                //xlWorkSheet.Cells[1, 9] = "AGENCIA";
                xlWorkSheet.Cells[1, 9] = "ESTADO";
                xlWorkSheet.Cells[1, 10] = "ENTRADA_ACI";
                //xlWorkSheet.Cells[1, 12] = "ENTRADA_IMSERSO";
                //xlWorkSheet.Cells[1, 13] = "Solicitud";
                //xlWorkSheet.Cells[1, 11] = "DIAS_ESTANCIA";
                //xlWorkSheet.Cells[1, 15] = "TURNO_IMSERSO";
                xlWorkSheet.Cells[1, 11] = "Diferente_NIF_Titular";
                xlWorkSheet.Cells[1, 12] = "Plazas_Diferentes";
                xlWorkSheet.Cells[1, 13] = "Reserva_Anulada";
                xlWorkSheet.Cells[1, 14] = "Dias_Estancia_Diferentes";
                xlWorkSheet.Cells[1, 15] = "Diferente_NIF_Acompañante";
                //xlWorkSheet.Cells[1, 17] = "Fecha_Entrada_Diferente";
                //xlWorkSheet.Cells[1, 20] = "Prereserva_Estado_No";
                xlWorkSheet.Cells[1, 16] = "Comprobar_Nombre_Apellidos_Titular";
                xlWorkSheet.Cells[1, 17] = "Fecha_Nacimiento_Erronea_Titular";
                xlWorkSheet.Cells[1, 18] = "Comprobar_Nombre_Apellidos_Acompanate";
                xlWorkSheet.Cells[1, 19] = "Fecha_Nacimiento_Erronea_Acompanate";
                xlWorkSheet.Cells[1, 20] = "Balneario";
                xlWorkSheet.Cells[1, 21] = "Hotel";

                // xlWorkSheet.Cells[1, 17] = "OBSERVACION";

                xlWorkBook.SaveAs(curFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }




            tabla_existe_registro_contratos = this.misfuncionesBDExpedientes.obtenerDatable("SELECT  sietedias FROM contratos_imserso WHERE sietedias <>''  GROUP BY  sietedias ");





            int contador_contratos = 0;
            foreach (DataRow dr in tabla_existe_registro_contratos.Rows)
            {

                if (contador_contratos == 0)
                {


                    if (this.checkBox3.Checked)
                    {
                        if (Convert.ToString(dr["sietedias"]) != "")
                        {
                            contratos_imsesro = "'" + Convert.ToString(dr["sietedias"]) + "'";
                        }
                    }
                    else
                    {

                        contratos_imsesro = "'" + Convert.ToString(dr["oncedias"]) + "'";
                        contratos_imsesro += ",'" + Convert.ToString(dr["nuevedias"]) + "'";


                    }



                    contador_contratos++;


                }
                else
                {



                    if (this.checkBox3.Checked)
                    {
                        if (Convert.ToString(dr["sietedias"]) != "" && contratos_imsesro != "")
                        {
                            contratos_imsesro += ",'" + Convert.ToString(dr["sietedias"]) + "'";

                        }



                    }
                    else
                    {

                        contratos_imsesro += ",'" + Convert.ToString(dr["oncedias"]) + "'";
                        contratos_imsesro += ",'" + Convert.ToString(dr["nuevedias"]) + "'";


                    }





                }


            }





            // string curFile = textBox1.Text + @"\Advertencias" + mesfichero + "_" + anyofichero + "\\AvertenciasClientes_" + fecha_actual_strg + ".xls";



            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
            string sql = null;
            MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + curFile + ";Extended Properties=Excel 12.0;");
            MyConnection.Open();
            myCommand.Connection = MyConnection;




            //Reservas

            bool entro_reserva_cliente = false;
            string sql_r = "";


            contratos_imsesro = "'" + System.Configuration.ConfigurationSettings.AppSettings["contrato_alta_tv"] + "','" + System.Configuration.ConfigurationSettings.AppSettings["contrato_baja_tv"] + "'";
            sql_r = "SELECT res_DR_bon_str,max(res_est_int) as res_est_int, hue_nif_str,hue_des_str,MAX(res_ent_dat) as res_ent_dat,max(RES_GUID) as RES_GUID,max(res_spe_adu_int) as res_spe_adu_int ,age_nom_str,DATEDIFF(day, max(res_ent_dat), max(res_sal_dat)) as dias,edi_des_str,Edificios.EDI_GUID,res_DR_bon_str    FROM Reservas INNER JOIN HuespedK ON HuespedK.HUE_GUID=Reservas.HUE_GUID INNER JOIN Huespedes ON Huespedes.HUE_GUID=Reservas.HUE_GUID  INNER JOIN contratos ON contratos.CTA_GUID =Reservas.CTA_GUID INNER JOIN AGENCIAs ON AGENCIAs.AGE_GUID = Reservas.AGE_GUID LEFT JOIN Edificios ON Edificios.EDI_GUID  = Reservas.EDI_GUID  where hue_nif_str IS NOT NULL AND hue_nif_str <> '' AND YEAR(res_ent_dat) =" + year_fichero + "  AND res_est_int IN (0,2)  AND Reservas.AGE_GUID='" + this.agencia_balneario + "' AND Reservas.EDI_GUID IN (" + this.EDI_GUID + ")  AND  RES_GUID NOT IN ( SELECT AAR_GUID FROM AdvertenciaAsignadaReservas WHERE ADV_GUID=43 ) GROUP BY hue_nif_str,age_nom_str,hue_des_str,edi_des_str,Edificios.EDI_GUID,res_DR_bon_str  ";



            DataTable tabla_existe_registro = new DataTable();



            /*Obtenemos el id del ultimo contador*/
            tabla_existe_registro = this.misfuncionesBD.obtenerDatable(sql_r);




            string TIPO_ADVERTENCIA = "";
            string estado = "";

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {




                fichero_advertenciasclientes = "OJO ----> Exite una Reserva " + Convert.ToString(dr["RES_GUID"]) + " --> " + Convert.ToString(dr["hue_nif_str"]) + "  " + Convert.ToString(dr["hue_des_str"]);
                entro_reserva_cliente = true;

                switch (Convert.ToString(dr["res_est_int"]))
                {
                    case "0":
                        estado = "RESERVA";

                        break;
                    case "2":
                        estado = "Checkin/Checkout";
                        break;
                    case "4":
                        estado = "ANUALADA";
                        break;
                }

                //TIPO_ADVERTENCIA = "Existe Reserva";
                //sql = "Insert into [Hoja1$] (" + CAMPOS_excel_fichero + ") values('" + Convert.ToString(dr["hue_nif_str"]) + "','" + Convert.ToString(dr["hue_des_str"]) + "','RESERVA','" + Convert.ToString(dr["RES_GUID"]) + "','" + Convert.ToString(dr["res_spe_adu_int"]) + "','" + Convert.ToString(dr["age_nom_str"]) + "','" + Convert.ToString(estado) + "','" + Convert.ToDateTime(dr["res_ent_dat"]).ToString("dd/MM/yyyy") + "','" + Convert.ToString(dr["dias"]) + "'";
                //sql = "Insert into [Hoja1$] (" + CAMPOS_excel_fichero + ") values('" + Convert.ToString(dr["hue_nif_str"]) + "','','','','','RESERVA','" + Convert.ToString(dr["RES_GUID"]) + "','" + Convert.ToString(dr["res_spe_adu_int"]) + "','" + Convert.ToString(dr["age_nom_str"]) + "','" + Convert.ToString(estado) + "','" + Convert.ToDateTime(dr["res_ent_dat"]).ToString("dd/MM/yyyy") + "','','','" + Convert.ToString(dr["dias"]) + "','','" + obtener_balneario_edificio(this.id_balneario_seleccionado.ToString(), Convert.ToString(dr["EDI_GUID"])) + "','" + Convert.ToString(dr["edi_des_str"]) + "'";
                sql = "Insert into [Hoja1$] (" + CAMPOS_excel_fichero + ") values('" + Convert.ToString(dr["hue_nif_str"]) + "','','','','" + Convert.ToString(dr["res_DR_bon_str"]) + "','RESERVA','" + Convert.ToString(dr["RES_GUID"]) + "','" + Convert.ToString(dr["res_spe_adu_int"]) + "','" + Convert.ToString(estado) + "','" + Convert.ToDateTime(dr["res_ent_dat"]).ToString("dd/MM/yyyy") + "','" + this.destino_FINAL + "','" + Convert.ToString(dr["edi_des_str"]) + "'";

                for (int i_h = 13; i_h < 22; i_h++)
                {


                    sql += ",''";

                }

                sql += ")";

                myCommand.CommandText = sql;
                myCommand.ExecuteNonQuery();





            }



        }

        public void HuespedesNoEstanImsersoExcel(string mes, string year_fichero, ArrayList ListaNifsImsersoMes)
        {


            string nif = "";
            string contratos_imsesro = "";
            ArrayList expediente_estado = new ArrayList();
            string no_alojados_advertencias = "";

            string mesfichero = Convert.ToString(mes);
            string anyofichero = Convert.ToString(year_fichero);



            CAMPOS_excel_fichero = "NifSol_ACI,NifSol_IMSERSO,NifAcom_ACI,NifAcom_IMSERSO,Expediente,Tipo,RESERVA,ADULTOS,AGENCIA,ESTADO,ENTRADA_ACI,ENTRADA_IMSERSO,Solicitud,DIAS_ESTANCIA,TURNO_IMSERSO,Balneario,Hotel,Diferente_NIF_Titular,Plazas_Diferentes,Reserva_Anulada,Dias_Estancia_Diferentes,Diferente_NIF_Acompañante,Fecha_Entrada_Diferente,Prereserva_Estado_No";

            //Obtenemos los contratos

            DataTable tabla_existe_registro_contratos = new DataTable();

            /*Obtenemos el id del ultimo contador*/




            if (!File.Exists(curFile))
            {


                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();


                curFile = textBox1.Text + @"\AdvertenciasImserso_" + mesfichero + "_" + anyofichero + "\\Advertencias_" + fecha_actual_strg + ".xls";

                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "NifSol_ACI";
                xlWorkSheet.Cells[1, 2] = "NifSol_IMSERSO";
                xlWorkSheet.Cells[1, 3] = "NifAcom_ACI";
                xlWorkSheet.Cells[1, 4] = "NifAcom_IMSERSO";
                xlWorkSheet.Cells[1, 5] = "Expediente";
                xlWorkSheet.Cells[1, 6] = "Tipo";
                xlWorkSheet.Cells[1, 7] = "RESERVA";
                xlWorkSheet.Cells[1, 8] = "ADULTOS";
                xlWorkSheet.Cells[1, 9] = "AGENCIA";
                xlWorkSheet.Cells[1, 10] = "ESTADO";
                xlWorkSheet.Cells[1, 11] = "ENTRADA_ACI";
                xlWorkSheet.Cells[1, 12] = "ENTRADA_IMSERSO";
                xlWorkSheet.Cells[1, 13] = "Solicitud";
                xlWorkSheet.Cells[1, 14] = "DIAS_ESTANCIA";
                xlWorkSheet.Cells[1, 15] = "TURNO_IMSERSO";
                xlWorkSheet.Cells[1, 16] = "Diferente_NIF_Titular";
                xlWorkSheet.Cells[1, 17] = "Plazas_Diferentes";
                xlWorkSheet.Cells[1, 18] = "Reserva_Anulada";
                xlWorkSheet.Cells[1, 19] = "Dias_Estancia_Diferentes";
                xlWorkSheet.Cells[1, 20] = "Diferente_NIF_Acompañante";
                xlWorkSheet.Cells[1, 21] = "Fecha_Entrada_Diferente";
                xlWorkSheet.Cells[1, 22] = "Prereserva_Estado_No";
                xlWorkSheet.Cells[1, 23] = "Balneario";
                xlWorkSheet.Cells[1, 24] = "Hotel";

                // xlWorkSheet.Cells[1, 17] = "OBSERVACION";

                xlWorkBook.SaveAs(curFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }



            if (this.checkBox3.Checked)
            {
                tabla_existe_registro_contratos = this.misfuncionesBDExpedientes.obtenerDatable("SELECT  sietedias FROM contratos_imserso WHERE sietedias <>''  GROUP BY  sietedias ");
            }
            else
            {

                tabla_existe_registro_contratos = this.misfuncionesBDExpedientes.obtenerDatable("SELECT  oncedias, nuevedias FROM contratos_imserso GROUP BY oncedias, nuevedias ");
            }




            int contador_contratos = 0;
            foreach (DataRow dr in tabla_existe_registro_contratos.Rows)
            {

                if (contador_contratos == 0)
                {


                    if (this.checkBox3.Checked)
                    {
                        if (Convert.ToString(dr["sietedias"]) != "")
                        {
                            contratos_imsesro = "'" + Convert.ToString(dr["sietedias"]) + "'";
                        }
                    }
                    else
                    {

                        contratos_imsesro = "'" + Convert.ToString(dr["oncedias"]) + "'";
                        contratos_imsesro += ",'" + Convert.ToString(dr["nuevedias"]) + "'";


                    }



                    contador_contratos++;


                }
                else
                {



                    if (this.checkBox3.Checked)
                    {
                        if (Convert.ToString(dr["sietedias"]) != "" && contratos_imsesro != "")
                        {
                            contratos_imsesro += ",'" + Convert.ToString(dr["sietedias"]) + "'";

                        }



                    }
                    else
                    {

                        contratos_imsesro += ",'" + Convert.ToString(dr["oncedias"]) + "'";
                        contratos_imsesro += ",'" + Convert.ToString(dr["nuevedias"]) + "'";


                    }





                }


            }





            // string curFile = textBox1.Text + @"\Advertencias" + mesfichero + "_" + anyofichero + "\\AvertenciasClientes_" + fecha_actual_strg + ".xls";



            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
            string sql = null;
            MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + curFile + ";Extended Properties=Excel 12.0;");
            MyConnection.Open();
            myCommand.Connection = MyConnection;




            //Reservas

            bool entro_reserva_cliente = false;
            string sql_r = "";


            if (this.checkBox3.Checked)
            {


                contratos_imsesro = "'" + System.Configuration.ConfigurationSettings.AppSettings["contrato_alta_tv"] + "','" + System.Configuration.ConfigurationSettings.AppSettings["contrato_baja_tv"] + "'";
                sql_r = "SELECT max(res_est_int) as res_est_int, hue_nif_str,hue_des_str,MAX(res_ent_dat) as res_ent_dat,max(RES_GUID) as RES_GUID,max(res_spe_adu_int) as res_spe_adu_int ,age_nom_str,DATEDIFF(day, max(res_ent_dat), max(res_sal_dat)) as dias,edi_des_str,Edificios.EDI_GUID   FROM Reservas INNER JOIN HuespedK ON HuespedK.HUE_GUID=Reservas.HUE_GUID INNER JOIN Huespedes ON Huespedes.HUE_GUID=Reservas.HUE_GUID  INNER JOIN contratos ON contratos.CTA_GUID =Reservas.CTA_GUID INNER JOIN AGENCIAs ON AGENCIAs.AGE_GUID = Reservas.AGE_GUID LEFT JOIN Edificios ON Edificios.EDI_GUID  = Reservas.EDI_GUID  where hue_nif_str IS NOT NULL AND hue_nif_str <> '' AND YEAR(res_ent_dat) =" + year_fichero + "  AND res_est_int IN (0,2)  AND Reservas.AGE_GUID='" + this.agencia_balneario + "' AND Reservas.EDI_GUID IN (" + this.EDI_GUID + ")   GROUP BY hue_nif_str,age_nom_str,hue_des_str,edi_des_str,Edificios.EDI_GUID ";


            }
            else
            {


                sql_r = "SELECT max(res_est_int) as res_est_int,hue_nif_str,hue_des_str,MAX(res_ent_dat) as res_ent_dat,max(RES_GUID) as RES_GUID,max(res_spe_adu_int) as res_spe_adu_int ,age_nom_str,DATEDIFF(day, max(res_ent_dat), max(res_sal_dat)) as dias,edi_des_str,Edificios.EDI_GUID  FROM Reservas INNER JOIN HuespedK ON HuespedK.HUE_GUID=Reservas.HUE_GUID INNER JOIN Huespedes ON Huespedes.HUE_GUID=Reservas.HUE_GUID  INNER JOIN contratos ON contratos.CTA_GUID =Reservas.CTA_GUID  INNER JOIN AGENCIAs ON AGENCIAs.AGE_GUID = Reservas.AGE_GUID LEFT JOIN Edificios ON Edificios.EDI_GUID  = Reservas.EDI_GUID  where hue_nif_str IS NOT NULL AND hue_nif_str <> '' AND  YEAR(res_ent_dat) =" + year_fichero + " AND MONTH(res_ent_dat) =" + mes + "   AND res_est_int IN (0,2)  AND Reservas.AGE_GUID='" + this.agencia_balneario + "' AND Reservas.EDI_GUID IN (" + this.EDI_GUID + ")  GROUP BY hue_nif_str,age_nom_str,hue_des_str,edi_des_str,Edificios.EDI_GUID ";

            }


            DataTable tabla_existe_registro = new DataTable();



            /*Obtenemos el id del ultimo contador*/
            tabla_existe_registro = this.misfuncionesBD.obtenerDatable(sql_r);




            string TIPO_ADVERTENCIA = "";
            string estado = "";

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {

                if (nif_fichero_imserso_mes.Contains(formatearnif(Convert.ToString(dr["hue_nif_str"]))) == false)
                {


                    fichero_advertenciasclientes = "OJO ----> Exite una Reserva " + Convert.ToString(dr["RES_GUID"]) + " --> " + Convert.ToString(dr["hue_nif_str"]) + "  " + Convert.ToString(dr["hue_des_str"]);
                    entro_reserva_cliente = true;

                    switch (Convert.ToString(dr["res_est_int"]))
                    {
                        case "0":
                            estado = "RESERVA";

                            break;
                        case "2":
                            estado = "Checkin/Checkout";
                            break;
                        case "4":
                            estado = "ANUALADA";
                            break;
                    }

                    //TIPO_ADVERTENCIA = "Existe Reserva";
                    //sql = "Insert into [Hoja1$] (" + CAMPOS_excel_fichero + ") values('" + Convert.ToString(dr["hue_nif_str"]) + "','" + Convert.ToString(dr["hue_des_str"]) + "','RESERVA','" + Convert.ToString(dr["RES_GUID"]) + "','" + Convert.ToString(dr["res_spe_adu_int"]) + "','" + Convert.ToString(dr["age_nom_str"]) + "','" + Convert.ToString(estado) + "','" + Convert.ToDateTime(dr["res_ent_dat"]).ToString("dd/MM/yyyy") + "','" + Convert.ToString(dr["dias"]) + "'";
                    sql = "Insert into [Hoja1$] (" + CAMPOS_excel_fichero + ") values('" + Convert.ToString(dr["hue_nif_str"]) + "','','','','','RESERVA','" + Convert.ToString(dr["RES_GUID"]) + "','" + Convert.ToString(dr["res_spe_adu_int"]) + "','" + Convert.ToString(dr["age_nom_str"]) + "','" + Convert.ToString(estado) + "','" + Convert.ToDateTime(dr["res_ent_dat"]).ToString("dd/MM/yyyy") + "','','','" + Convert.ToString(dr["dias"]) + "','','" + obtener_balneario_edificio(this.id_balneario_seleccionado.ToString(), Convert.ToString(dr["EDI_GUID"])) + "','" + Convert.ToString(dr["edi_des_str"]) + "'";


                    for (int i_h = 16; i_h < 23; i_h++)
                    {


                        sql += ",''";

                    }

                    sql += ")";

                    myCommand.CommandText = sql;
                    myCommand.ExecuteNonQuery();


                }


            }



        }

      








        public string obtnener_balneario_conexion(string id_balneario)
        {

            string balne_conexion = "";


            DataTable tabla_existe_registro = new DataTable();
            string id_estado = "";
            /*Obtenemos el id del ultimo contador*/
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT [id],[balneario] ,[destino],[conexion] ,[agencia],[tipo_habitacion] FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                balne_conexion = Convert.ToString(dr["conexion"]);


            }



            this.conexion_balneario = balne_conexion;
            return balne_conexion;

        }



        public string obtener_estado_bd(string estado)
        {



            DataTable tabla_existe_registro = new DataTable();
            string id_estado = "";
            /*Obtenemos el id del ultimo contador*/
            tabla_existe_registro = this.misfuncionesBD.obtenerDatable("SELECT   epr_GUID, epr_COD_str, epr_des_str, epr_color_lng, epr_color_text_lng, epr_def_bln, epr_cal_bln, epr_anu_bln FROM PreReservaEstados WHERE epr_des_str='" + estado + "' ");
            // fichero_advertenciasclientes += "SELECT hue_nif_str,res_ent_dat,RES_GUID,CASE res_est_int WHEN 1 THEN 'Transpaso AI' WHEN 2 THEN 'SI' WHEN 3 THEN 'Pendiente' WHEN 4 THEN 'EQUIVOCADO' WHEN 5 THEN 'NO LOCALIZADO' WHEN 6 THEN 'VOLVER A LLAMAR' WHEN 7 THEN 'PROPIO' WHEN 8 THEN 'PROPIO (Otras opciones)' WHEN 9 THEN 'NO' ELSE '' END AS Estadif (localIP.Contains("192.168.0."))

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                id_estado = Convert.ToString(dr["epr_GUID"]);


            }

            return id_estado;

        }



        /*obtner tipo de pension*/
        public string obtener_pension(string id_balneario)
        {



            string pension = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT PEN_GUID FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                pension = Convert.ToString(dr["PEN_GUID"]);


            }


            return pension;

        }


        public string obtener_tipo_habitacion(string id_balneario)
        {



            string tipo_habitacion = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT [id],[balneario] ,[destino],[conexion] ,[agencia],[tipo_habitacion] FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                tipo_habitacion = Convert.ToString(dr["tipo_habitacion"]);


            }


            return tipo_habitacion;

        }


        public string obtener_tipo_edificio(string id_balneario)
        {



            string tipo_edificio = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT [id],[balneario] ,[destino],[conexion] ,[agencia],[tipo_habitacion], edificio FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                tipo_edificio = Convert.ToString(dr["edificio"]);


            }


            return tipo_edificio;

        }




        public string obtener_balneario_edificio(string id_balneario, string EDI_GUID)
        {



            string balneario = "";





            DataTable tabla_existe_registro = new DataTable();
            string sql = "SELECT [destino]  FROM [config_balneario_prereservas] INNER JOIN edificios_balnearios ON edificios_balnearios.id_balneario=config_balneario_prereservas.id WHERE id_balneario='" + id_balneario + "' AND edificios_balnearios.EDI_GUID='" + EDI_GUID + "' ORDER BY balneario ASC";
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable(sql);





            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                balneario = Convert.ToString(dr["destino"]);


            }


            return balneario;

        }


        public string obtener_agencia(string id_balneario)
        {



            string agencia = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT [id],[balneario] ,[destino],[conexion] ,[agencia],[tipo_habitacion], agencia_generalitat FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {

                if (this.checkBox3.Checked == true)
                {

                    agencia = Convert.ToString(dr["agencia_generalitat"]);

                }
                else
                {


                    agencia = Convert.ToString(dr["agencia"]);


                }



            }


            return agencia;

        }




        public string obtener_epr_GUID_tv(string id_balneario)
        {



            string epr_GUID_tv = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT epr_GUID_tv FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                epr_GUID_tv = Convert.ToString(dr["epr_GUID_tv"]);






            }


            return epr_GUID_tv;

        }

        public string obtener_formatodatetime(string id_balneario)
        {



            string formato = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT [id],[balneario] ,[destino],[conexion] ,[agencia],[tipo_habitacion], formatodatetime FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                formato = Convert.ToString(dr["formatodatetime"]);


            }


            return formato;

        }

        [Obsolete]
        public void llenar_combo_balnearios()
        {

            string hostName = Dns.GetHostName(); 
            string myIP = Dns.GetHostByName(hostName).AddressList[0].ToString();
          


            DataTable tabla_existe_combo = new DataTable();
            String sql = "SELECT [id],[balneario] ,[destino],[conexion] ,[agencia],[tipo_habitacion] FROM [config_balneario_prereservas] WHERE id IN (#ids#) ORDER BY [id] ASC";


            sql = sql.Replace("#ids#", ConfigurationManager.AppSettings["balnearios"]);


            tabla_existe_combo = this.misfuncionesBDExpedientes.obtenerDatable(sql);



            this.comboBox1.DataSource = tabla_existe_combo;
            this.comboBox1.ValueMember = "id";
            this.comboBox1.DisplayMember = "destino";


        }




        //Obtner prereservas pendientes

        public void ObtenerPrereservasPendientes(string mes=null)
        {

            //MessageBox.Show("hola");
            string sql = "SELECT res_fec_dat,res_DR_obs_str,PresReservas.HUE_GUID,res_spe_adu_int,res_ent_dat, epr_des_str,RES_GUID,res_sal_dat,hue_des_str FROM PresReservas  INNER JOIN PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID  INNER JOIN Huespedes ON Huespedes.hue_GUID=PresReservas.hue_GUID WHERE epr_des_str='Pendiente' AND YEAR(res_ent_dat)=YEAR(getdate()) ";

            if (mes != null)
            {
                sql += "  AND MONTH(res_ent_dat)=" + mes;
            }
            else {

                sql += "  AND MONTH(res_ent_dat)=MONTH(getdate())";

            }
            

            DataTable datetablependientes = new DataTable();

            datetablependientes = this.misfuncionesBD.obtenerDatable(sql);
            //tetablependientes = this.misfuncionesBDExpedientes.obtenerDatable("SELECT * FROM [Localidades] ");
            //dataGridView1.DataSource = this.misfuncionesBD.obtenerDataSet(sql);
            //dataGridView2.DataSource=this.misfuncionesBDExpedientes.obtenerDatable();
            //dataGridView2.DataMember = "Localidad_OK";


            int contador_lineas_gid = 1;

            foreach (DataRow dr in datetablependientes.Rows)
            {
                int renglon = dataGridView2.Rows.Add();
                dataGridView2.Rows[renglon].Cells["contador2"].Value = Convert.ToString(contador_lineas_gid);
                dataGridView2.Rows[renglon].Cells["entrada2"].Value = Convert.ToDateTime(dr["res_ent_dat"]).ToString("dd/MM/yyyy");
                dataGridView2.Rows[renglon].Cells["salida2"].Value = Convert.ToDateTime(dr["res_sal_dat"]).ToString("dd/MM/yyyy");
                dataGridView2.Rows[renglon].Cells["estado2"].Value = Convert.ToString(dr["epr_des_str"]);
                dataGridView2.Rows[renglon].Cells["reserva2"].Value = Convert.ToString(dr["RES_GUID"]);
                dataGridView2.Rows[renglon].Cells["huesped2"].Value = Convert.ToString(dr["hue_des_str"]);
                dataGridView2.Rows[renglon].Cells["plazas2"].Value = Convert.ToString(dr["res_spe_adu_int"]);
                dataGridView2.Rows[renglon].Cells["alta2"].Value = Convert.ToDateTime(dr["res_fec_dat"]).ToString("dd/MM/yyyy");
                dataGridView2.Rows[renglon].Cells["observaciones2"].Value = Convert.ToString(dr["res_DR_obs_str"]);
                dataGridView2.Rows[renglon].Cells["codigo_huesped2"].Value = Convert.ToString(dr["HUE_GUID"]);
                contador_lineas_gid++;

            }

        }


        public void PrereservasPendientesIntersatadas()
        {

            //MessageBox.Show("hola");
            string sql = "";
            string prereserva_inser = "";

            for (int i = 0; i <= this.PrereservasInsertadas.Count - 1; i++)
            {
                if (i == this.PrereservasInsertadas.Count - 1)
                {
                    prereserva_inser += Convert.ToString(PrereservasInsertadas[i]);
                }
                else
                {

                    prereserva_inser += Convert.ToString(PrereservasInsertadas[i]) + ",";

                }

            }

            sql = "SELECT res_fec_dat,res_DR_obs_str,PresReservas.HUE_GUID,res_spe_adu_int,res_ent_dat, epr_des_str,RES_GUID,res_sal_dat,hue_des_str FROM PresReservas  INNER JOIN PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID  INNER JOIN Huespedes ON Huespedes.hue_GUID=PresReservas.hue_GUID WHERE RES_GUID IN (" + prereserva_inser + ")";
            DataTable datetablependientes = new DataTable();
            datetablependientes = this.misfuncionesBD.obtenerDatable(sql);




            //tetablependientes = this.misfuncionesBDExpedientes.obtenerDatable("SELECT * FROM [Localidades] ");
            //dataGridView1.DataSource = this.misfuncionesBD.obtenerDataSet(sql);
            //dataGridView2.DataSource=this.misfuncionesBDExpedientes.obtenerDatable();
            //dataGridView2.DataMember = "Localidad_OK";


            int contador_lineas_gid = 1;

            foreach (DataRow dr in datetablependientes.Rows)
            {
                int renglon = dataGridView1.Rows.Add();
                dataGridView1.Rows[renglon].Cells["contador1"].Value = Convert.ToString(contador_lineas_gid);
                dataGridView1.Rows[renglon].Cells["entrada1"].Value = Convert.ToDateTime(dr["res_ent_dat"]).ToString("dd/MM/yyyy");
                dataGridView1.Rows[renglon].Cells["salida1"].Value = Convert.ToDateTime(dr["res_sal_dat"]).ToString("dd/MM/yyyy");
                dataGridView1.Rows[renglon].Cells["estado1"].Value = Convert.ToString(dr["epr_des_str"]);
                dataGridView1.Rows[renglon].Cells["reserva1"].Value = Convert.ToString(dr["RES_GUID"]);
                dataGridView1.Rows[renglon].Cells["huesped1"].Value = Convert.ToString(dr["hue_des_str"]);
                dataGridView1.Rows[renglon].Cells["plazas1"].Value = Convert.ToString(dr["res_spe_adu_int"]);
                dataGridView1.Rows[renglon].Cells["alta1"].Value = Convert.ToDateTime(dr["res_fec_dat"]).ToString("dd/MM/yyyy");
                dataGridView1.Rows[renglon].Cells["codigo_huesped1"].Value = Convert.ToString(dr["HUE_GUID"]);
                dataGridView1.Rows[renglon].Cells["observaciones1"].Value = Convert.ToString(dr["res_DR_obs_str"]);
                contador_lineas_gid++;
            }

        }

        public void ObtenerPrereservasPendientesRepetidas(string mes)
        {

            //MessageBox.Show("hola");
            string sql = "SELECT hue_guid,count(*) FROM [PresReservas]  WHERE hue_guid<>'' AND  MONTH(res_ent_dat)='" + mes + "' GROUP BY hue_guid HAVING count(*) > 1 ";
            ArrayList arraylistReptidas = new ArrayList();

            DataTable datetablependientes = new DataTable();

            datetablependientes = this.misfuncionesBD.obtenerDatable(sql);
            string repetidas_pre = "";


            foreach (DataRow dr in datetablependientes.Rows)
            {
                arraylistReptidas.Add(Convert.ToString(dr["hue_guid"]));
            }

            if (arraylistReptidas.Count > 0)
            {
                for (int i = 0; i <= arraylistReptidas.Count - 1; i++)
                {
                    if (i == arraylistReptidas.Count - 1)
                    {



                        repetidas_pre += Convert.ToString(arraylistReptidas[i]);


                    }
                    else
                    {




                        repetidas_pre += Convert.ToString(arraylistReptidas[i]) + ",";
                    }

                }



                sql = "SELECT  res_fec_dat,res_DR_obs_str,PresReservas.HUE_GUID,res_spe_adu_int,res_ent_dat, epr_des_str,RES_GUID,res_sal_dat,hue_des_str FROM PresReservas  INNER JOIN PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID  INNER JOIN Huespedes ON Huespedes.hue_GUID=PresReservas.hue_GUID WHERE  YEAR(res_ent_dat)=YEAR(getdate()) AND PresReservas.hue_GUID IN (" + repetidas_pre + ") AND MONTH(res_ent_dat)='" + mes + "' ORDER BY PresReservas.hue_GUID ASC ";
                datetablependientes = this.misfuncionesBD.obtenerDatable(sql);

                int contador_lineas_gid = 1;
                foreach (DataRow dr in datetablependientes.Rows)
                {
                    int renglon = dataGridView3.Rows.Add();
                    dataGridView3.Rows[renglon].Cells["contador3"].Value = Convert.ToString(contador_lineas_gid);
                    dataGridView3.Rows[renglon].Cells["entrada3"].Value = Convert.ToDateTime(dr["res_ent_dat"]).ToString("dd/MM/yyyy");
                    dataGridView3.Rows[renglon].Cells["salida3"].Value = Convert.ToDateTime(dr["res_sal_dat"]).ToString("dd/MM/yyyy");
                    dataGridView3.Rows[renglon].Cells["estado3"].Value = Convert.ToString(dr["epr_des_str"]);
                    dataGridView3.Rows[renglon].Cells["reserva3"].Value = Convert.ToString(dr["RES_GUID"]);
                    dataGridView3.Rows[renglon].Cells["huesped3"].Value = Convert.ToString(dr["hue_des_str"]);
                    dataGridView3.Rows[renglon].Cells["plazas3"].Value = Convert.ToString(dr["res_spe_adu_int"]);
                    dataGridView3.Rows[renglon].Cells["alta3"].Value = Convert.ToDateTime(dr["res_fec_dat"]).ToString("dd/MM/yyyy");
                    dataGridView3.Rows[renglon].Cells["observaciones3"].Value = Convert.ToString(dr["res_DR_obs_str"]);
                    dataGridView3.Rows[renglon].Cells["codigo_huesped"].Value = Convert.ToString(dr["HUE_GUID"]);

                    contador_lineas_gid++;

                }

            }




        }


        public void ActulizarHuespeded(int hue_guid, string hue_des_str, string hue_dom_str, string hue_pob_str, string hue_cp_str, string hue_nif_str, string hue_prv_str, string telefono1, string telefono2, string fecha_turno, string tipo_nif, string nifsolictente, string birthdate)
        {

            int HUE_GUID = 0;
            int Con_cod_lng = 0;
            string campos = "";
            string values = "";
            DateTime fecha_acutal = DateTime.Now;


            //MessageBox.Show("hola");
            string sql = "SELECT huespedes.hue_guid,hue_dom_str,hue_pob_str,hue_cp_str,hue_mai_str,hue_prv_str FROM [huespedes] INNER JOIN HuespedK ON HuespedK.hue_guid=huespedes.hue_guid WHERE HuespedK.hue_guid='" + Convert.ToString(hue_guid) + "'";
            DataTable datetablhuesped = new DataTable();

            datetablhuesped = this.misfuncionesBD.obtenerDatable(sql);
            string campos_actulizar = "";
            bool entro_actulizar = false;



            foreach (DataRow dr in datetablhuesped.Rows)
            {

                campos_actulizar += "hue_pul_str='ActulizacionImserso'";


                if (Convert.ToString(dr["hue_dom_str"]).Length <= 2)
                {

                    campos_actulizar += ",hue_dom_str='" + hue_dom_str.Replace("'", "''") + "'";
                    entro_actulizar = true;

                }


                if (Convert.ToString(dr["hue_cp_str"]).Length <= 2)
                {


                    campos_actulizar += ",hue_cp_str='" + hue_cp_str + "'";
                    entro_actulizar = true;



                }



                if (Convert.ToString(dr["hue_prv_str"]).Length <= 2)
                {


                    campos_actulizar += ",hue_prv_str='" + hue_prv_str.Replace("'", "''") + "'";
                    entro_actulizar = true;

                }






                if (Convert.ToString(dr["hue_pob_str"]).Length <= 2)
                {


                    campos_actulizar += ",hue_pob_str='" + hue_pob_str.Replace("'", "''") + "'";
                    entro_actulizar = true;




                }




                if (entro_actulizar == true)
                {

                    //MessageBox.Show(campos_actulizar);
                    this.misfuncionesBD.update("HuespedK", campos_actulizar, " HUE_GUID='" + Convert.ToString(hue_guid) + "'");

                }





            }

            /* DataTable mitable = new DataTable();

             ArrayList ArrylistContact = new ArrayList();
             string sql_c = "SELECT first_name,last_name,primary_address_street,primary_address_city,primary_address_state,primary_address_postalcode,birthdate FROM  contacts WHERE id_nif='" + hue_nif_str +"'";
             MySqlDataAdapter mda = new MySqlDataAdapter(sql_c, this.msqlConnection);
             DataSet ds = new DataSet();
             mda.Fill(ds);
             mitable = ds.Tables[0];
             entro_actulizar = false;
             foreach (System.Data.DataRow dr in mitable.Rows)
             {




                 if (Convert.ToString(dr["birthdate"]).Length <= 2)
                 {


                     campos_actulizar = "birthdate='" + birthdate + "'";
                     entro_actulizar = true;




                 }

                 if (entro_actulizar == true)
                 {

                     MySqlCommand cmd = new MySqlCommand();

                     cmd.Connection = this.msqlConnection;
                     string sql_up = "UPDATE contacts SET birthdate = '" + birthdate + "' WHERE id_nif = '" + hue_nif_str + "'";
                     cmd.CommandText = sql_up;
                     //int numRowsUpdated = cmd.ExecuteNonQuery(); 

                 }




             }*/




        }

        private void button4_Click(object sender, EventArgs e)
        {

            string sql = "SELECT  epr_GUID, epr_COD_str, epr_des_str, epr_color_lng, epr_color_text_lng, epr_def_bln, epr_cal_bln, epr_anu_bln FROM PreReservaEstados WHERE epr_des_str='No' ";
            DataTable datetableestado = this.misfuncionesBD.obtenerDatable(sql);
            string id_estado = "";

            foreach (DataRow dr in datetableestado.Rows)
            {

                id_estado = Convert.ToString(dr["epr_GUID"]);
            }


            DialogResult result = MessageBox.Show("¿Estas seguro cambiar el estado?", "Estado", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {

                foreach (DataGridViewRow row in dataGridView3.Rows)
                {



                    if (Convert.ToBoolean(row.Cells["selecionar"].Value) == true)
                    {
                        this.misfuncionesBD.update("PresReservas", "epr_GUID='" + id_estado + "',res_DR_obs_str=res_DR_obs_str +'---'+'" + textBox2.Text + "'", " RES_GUID='" + Convert.ToString(row.Cells["reserva3"].Value) + "'");
                    }


                }

            }

            MessageBox.Show("Cambios Realizados");
            //this.dataGridView3.Dispose();
            //ObtenerPrereservasPendientesRepetidas(this.mes_del_fichero);




        }


        


        public string valortexto_balneario(int balnearo)
        {
            DataTable tabla_existe_combo_balneario = new DataTable();
            tabla_existe_combo_balneario = this.misfuncionesBDExpedientes.obtenerDatable("SELECT destino_texto,[id],[balneario] ,[destino],[conexion] ,[agencia],[tipo_habitacion] FROM [config_balneario_prereservas] WHERE id='" + Convert.ToString(balnearo) + "'");
            string text_balneario = "";



            foreach (DataRow dr in tabla_existe_combo_balneario.Rows)
            {

                text_balneario = Convert.ToString(dr["destino_texto"]);


            }


            return text_balneario;
        }


        public string obtener_codigo_contrato(string contrato)
        {
            DataTable tabla_Contratos = new DataTable();
            tabla_Contratos = this.misfuncionesBD.obtenerDatable("SELECT CTA_GUID FROM [Contratos] WHERE CTA_cod_str='" + contrato + "'");
            string contrato_id = "";



            foreach (DataRow dr in tabla_Contratos.Rows)
            {

                contrato_id = Convert.ToString(dr["CTA_GUID"]);


            }


            return contrato_id;
        }


        public string obtener_SPE_GUID(string id_balneario)
        {
            string tipo_SPE_GUID = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT SPE_GUID,SPE_GUID_TV FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {

                if (this.checkBox3.Checked)
                {

                    tipo_SPE_GUID = Convert.ToString(dr["SPE_GUID_TV"]);
                }
                else
                {


                    tipo_SPE_GUID = Convert.ToString(dr["SPE_GUID"]);
                }



            }


            return tipo_SPE_GUID;
        }




        public string obtener_EDI_GUID(string id_balneario)
        {
            string EDI_GUID = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT EDI_GUID FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {

                if (this.checkBox3.Checked)
                {

                    EDI_GUID = Convert.ToString(dr["EDI_GUID"]);
                }
                else
                {


                    EDI_GUID = Convert.ToString(dr["EDI_GUID"]);
                }



            }


            return EDI_GUID;
        }


        public string obtener_EDI_GUID_default(string id_balneario)
        {
            string EDI_GUID = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT EDI_GUID_default FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {

                if (this.checkBox3.Checked)
                {

                    EDI_GUID = Convert.ToString(dr["EDI_GUID_default"]);
                }
                else
                {


                    EDI_GUID = Convert.ToString(dr["EDI_GUID_default"]);
                }



            }


            return EDI_GUID;
        }



        //Obtner orgin reseserva

        public string obtener_ORR_GUID(string id_balneario, string origen)
        {


            string ORR_GUID = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT ORR_GUID_" + origen + " FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                ORR_GUID = Convert.ToString(dr["ORR_GUID_" + origen]);


            }


            return ORR_GUID;


        }



        public string obtener_advertencia(string id_balneario, string adv)
        {


            string adv_GUID = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT advertencia_" + adv + " FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {


                adv_GUID = Convert.ToString(dr["advertencia_" + adv]);


            }


            return adv_GUID;


        }


        public void PreReservasReservasAnulasOtraConvocatoria(string nif, string expediente)
        {


            char[] delimiters = new char[] { '-' };
            string rango_convocatorias;
            rango_convocatorias = ObtenerRangoConvocatoria(Convert.ToDateTime("2014-08-01"));
            string[] VectotConvocatorias = rango_convocatorias.Split(delimiters);


            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBD.obtenerDatable("SELECT hue_des_str,res_ent_dat,RES_GUID,PreReservaEstados.epr_des_str as Estado,prr_est_byt FROM PresReservas INNER JOIN  PreReservaEstados ON PreReservaEstados.epr_GUID=PresReservas.epr_GUID INNER JOIN HuespedK ON HuespedK.HUE_GUID=PresReservas.HUE_GUID INNER JOIN Huespedes ON Huespedes.HUE_GUID=PresReservas.HUE_GUID where (HuespedK.hue_nif_str='" + Convert.ToString(nif) + "' AND res_ent_dat between '" + Convert.ToDateTime(VectotConvocatorias[0]).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime(VectotConvocatorias[1]).ToString(formato_datetime) + "')");


            foreach (DataRow dr in tabla_existe_registro.Rows)
            {

                fichero_advertencias += "2º Convocatoria para este expediente " + expediente + " existe una PresReserva " + Convert.ToString(dr["RES_GUID"]) + " en estado " + Convert.ToString(dr["Estado"]) + " del NIF " + nif + Environment.NewLine + Environment.NewLine;

            }





            tabla_existe_registro = this.misfuncionesBD.obtenerDatable("SELECT   res_ent_dat,hue_des_str,hue_nif_str,res_ent_dat,RES_GUID,CASE res_est_int WHEN 0 THEN 'Reserva' WHEN 2 THEN 'Checkin/Checkout' WHEN 3 THEN 'No Show' WHEN 4 THEN 'Anulada' ELSE '' END AS Estado FROM Reservas INNER JOIN HuespedK ON HuespedK.HUE_GUID=Reservas.HUE_GUID INNER JOIN Huespedes ON Huespedes.HUE_GUID=Reservas.HUE_GUID  where (HuespedK.hue_nif_str='" + Convert.ToString(nif) + "' AND res_ent_dat between '" + Convert.ToDateTime(VectotConvocatorias[0]).ToString(formato_datetime) + "' AND   '" + Convert.ToDateTime(VectotConvocatorias[1]).ToString(formato_datetime) + "')");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {

                fichero_advertencias += "2º Convocatoria para este expediente " + expediente + " existe una Reserva " + Convert.ToString(dr["RES_GUID"]) + " en estado " + Convert.ToString(dr["Estado"]) + " del NIF " + nif + Environment.NewLine + Environment.NewLine;

            }





        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Cursor Files|*.rtf";
            openFileDialog1.Title = "Select a Cursor File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Assign the cursor in the Stream to the Form's Cursor property.
                textBox4.Text = openFileDialog1.FileName;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

            // If your RTF file isn't in the same folder as the .exe file for the project, 
            // specify the path to the file in the following assignment statement. 
            string path = textBox4.Text;
            string ruta = System.IO.Path.GetDirectoryName(path);


            string cabecera_fichero = leer_template(ruta);


            //MessageBox.Show(cabecera_fichero);



            if (File.Exists(ruta + @"\transformado.txt"))
            {
                File.Delete(ruta + @"\transformado.txt");
            }


            //Create the RichTextBox. (Requires a reference to System.Windows.Forms.)
            System.Windows.Forms.RichTextBox rtBox = new System.Windows.Forms.RichTextBox();

            // Get the contents of the RTF file. When the contents of the file are  
            // stored in the string (rtfText), the contents are encoded as UTF-16. 
            string rtfText = System.IO.File.ReadAllText(path);

            // Display the RTF text. This should look like the contents of your file.
            //System.Windows.Forms.MessageBox.Show(rtfText);

            // Use the RichTextBox to convert the RTF code to plain text.
            rtBox.Rtf = rtfText;
            string plainText = rtBox.Text;

            // Display the plain text in a MessageBox because the console can't  
            // display the Greek letters. You should see the following result: 
            //   The Greek word for "psyche" is spelled ψυχή. The Greek letters are
            //   encoded in Unicode.
            //   These characters are from the extended ASCII character set (Windows
            //   code page 1252): âäӑå
            //System.Windows.Forms.MessageBox.Show(plainText);
            int pos_fecha = plainText.IndexOf("Fecha Inicio: ");

            string fecha_inicio = plainText.Substring(pos_fecha + 13, 11);
            //fecha_inicio = fecha_inicio.Replace("-","/");
            string[] fecha_split = fecha_inicio.Split(new Char[] { '-' });
            fecha_inicio = fecha_split[2] + "-" + fecha_split[1] + "-" + fecha_split[0].Trim();


            fecha_inicio = String.Format("{0:yyyy-M-d}", fecha_inicio);
            int pos = plainText.IndexOf("DATOS DE LOS EXPEDIENTES");

            //MessageBox.Show(Convert.ToString(pos));


            string cabecera = "";

            string pruebas = plainText.Substring(pos + 24).Replace(" ______________________________________________________________", "#");

            pruebas = pruebas.Replace("Datos del Solicitante", " ");

            int pos_expediente = 0;
            int pos_n_expediente = 0;
            int pos_Expediente_Relacionado = 0;
            int pos_Tipo_de_Solicitud = 0;
            int pos_Plazas = 0;
            int pos_Confirmacion = 0;
            int pos_apelle_nombre = 0;
            int pos_dni = 0;
            int pos_telefono = 0;
            int pos_fecha_nacimiento = 0;
            int pos_Sexo = 0;
            int pos_direccion = 0;

            int pos_acompanyante = 0;





            int pos_apelle_nombre_acom = 0;
            int pos_dni_acom = 0;
            int pos_fecha_nacimiento_acom = 0;
            int pos_Sexo_acom = 0;




            string str_expediente = "    ";
            string str_n_expediente = "    ";
            string str_Expediente_Relacionado = "      ";
            string str_Tipo_de_Solicitud = "";
            string str_Plazas = "     ";
            string str_Confirmacion = "     ";
            string str_apelle_nombre = "     ";
            string str_dni = "      ";
            string str_telefono = "      ";
            string str_fecha_nacimiento = "     ";
            string str_Sexo = "     ";
            string str_direccion = "     ";
            string acomp = "     ";






            string cp = "";
            string poblacion = "";
            string str_direccion_calle = "";
            string[] direcion_final;



            string str_apelle_acom = "";
            string str_Sexo_acom = "";
            string str_dni_acom = "";
            string str_fecha_nacimiento_acom = "";


            string duracion = "A (7 DIAS)";
            string texto_final = "";


            string[] split_acom;


            string[] split = pruebas.Split(new Char[] { '#' });
            ArrayList myAL = new ArrayList();


            string expediente_final = "";

            for (int i = 1; i < split.Length; i++)
            {

                pos_expediente = split[i].IndexOf("Expediente:");
                pos_n_expediente = split[i].IndexOf("Nº Expediente:");
                pos_Expediente_Relacionado = split[i].IndexOf("Expediente Relacionado:");
                pos_Tipo_de_Solicitud = split[i].IndexOf("Tipo de Solicitud:");
                pos_Plazas = split[i].IndexOf("Plazas:");
                pos_Confirmacion = split[i].IndexOf("Confirmación:");
                pos_apelle_nombre = split[i].IndexOf("Apellidos y Nombre:");
                pos_dni = split[i].IndexOf("DNI:");
                pos_telefono = split[i].IndexOf("Teléfono:");
                pos_fecha_nacimiento = split[i].IndexOf("Fecha de Nacimiento:");
                pos_Sexo = split[i].IndexOf("Sexo:");
                pos_direccion = split[i].IndexOf("Dirección:");


                pos_acompanyante = split[i].IndexOf("Datos del Acompañante");



                str_expediente = split[i].Substring(pos_expediente, pos_n_expediente - 1).Replace("Expediente:", "").Trim();
                str_expediente = str_expediente.Replace(" ", "#");
                string[] split_expediene = str_expediente.Split(new Char[] { '#' });
                expediente_final = fecha_split[2].Trim() + "/" + split_expediene[0] + "G";

                str_n_expediente = split[i].Substring(pos_n_expediente, pos_Expediente_Relacionado - pos_n_expediente).Trim();
                str_n_expediente = str_n_expediente.Replace("Nº Expediente: ", "");




                str_Expediente_Relacionado = split[i].Substring(pos_Expediente_Relacionado, pos_Tipo_de_Solicitud - pos_Expediente_Relacionado).Trim();
                str_Expediente_Relacionado = str_Expediente_Relacionado.Replace("Expediente Relacionado: ", "");
                str_Tipo_de_Solicitud = split[i].Substring(pos_Tipo_de_Solicitud, pos_Plazas - pos_Tipo_de_Solicitud).Trim();
                str_Tipo_de_Solicitud = str_Tipo_de_Solicitud.Replace("Tipo de Solicitud:", "");
                str_Plazas = split[i].Substring(pos_Plazas, pos_Confirmacion - pos_Plazas).Trim();
                str_Plazas = str_Plazas.Replace("Plazas:", "");

                str_Confirmacion = split[i].Substring(pos_Confirmacion, pos_apelle_nombre - pos_Confirmacion).Trim();
                str_Confirmacion = str_Confirmacion.Replace("Confirmación:", "");


                str_apelle_nombre = split[i].Substring(pos_apelle_nombre, pos_dni - pos_apelle_nombre).Trim();
                str_apelle_nombre = str_apelle_nombre.Replace("Apellidos y Nombre:", "");

                str_dni = split[i].Substring(pos_dni, pos_telefono - pos_dni).Trim();
                str_dni = str_dni.Replace("DNI:", "");
                str_dni = formatearnif(CalculaNIF(str_dni));


                str_telefono = split[i].Substring(pos_telefono, pos_fecha_nacimiento - pos_telefono).Trim();
                str_telefono = str_telefono.Replace("Teléfono:", "");


                if (str_telefono.Trim() == "")
                {


                    str_telefono = "       ";

                }


                str_fecha_nacimiento = split[i].Substring(pos_fecha_nacimiento, pos_Sexo - pos_fecha_nacimiento).Trim();
                str_fecha_nacimiento = str_fecha_nacimiento.Replace("Fecha de Nacimiento:", "");



                str_Sexo = split[i].Substring(pos_Sexo, pos_direccion - pos_Sexo).Trim();
                str_Sexo = str_Sexo.Replace("Sexo:", "");

                //MessageBox.Show("ok");




                string dire = "";
                string[] split_direc;
                string[] split_direc_uno;
                string direc_final_dos;
                string[] split_direc_final;

                string datos_acomp = "";

                string direc_uno = "";


                string[] poblacion_split;

                cp = "";
                poblacion = "";
                str_direccion_calle = "";
                str_dni_acom = "         ";
                str_apelle_acom = "        ";

                if (pos_acompanyante > 0)
                {

                    str_direccion = split[i].Substring(pos_direccion).Trim();


                    acomp = str_direccion;
                    acomp = acomp.Replace("Datos del Acompañante", "#");
                    split_acom = acomp.Split(new Char[] { '#' });
                    str_direccion = split_acom[0].Replace("Dirección:", "");


                    dire = str_direccion.Replace("  - ", "#");
                    split_direc = dire.Split(new Char[] { '#' });

                    direc_final_dos = split_direc[0].Replace("    ", "#");
                    split_direc_uno = direc_final_dos.Split(new Char[] { '#' });

                    try
                    {
                        cp = split_direc_uno[1];
                    }
                    catch
                    {

                        cp = "00000";
                    }

                    str_direccion_calle = split_direc_uno[0];
                    poblacion = split_direc[1].Replace("\t", "");
                    poblacion = poblacion.Trim();



                    if (poblacion.IndexOf("/") > 0)
                    {

                        poblacion_split = poblacion.Split(new Char[] { '/' });
                        poblacion = poblacion_split[1];

                    }

                    poblacion = poblacion.Replace("'", "");

                    // MessageBox.Show(poblacion);
                    datos_acomp = split_acom[1];

                    pos_apelle_nombre_acom = datos_acomp.IndexOf("Apellidos y Nombre:");
                    pos_dni_acom = datos_acomp.IndexOf("DNI:");


                    pos_Sexo_acom = datos_acomp.IndexOf("Sexo:");
                    pos_fecha_nacimiento_acom = datos_acomp.IndexOf("Fecha de Nacimiento:");

                    str_apelle_acom = datos_acomp.Substring(pos_apelle_nombre_acom, pos_dni_acom - pos_apelle_nombre_acom).Trim();
                    str_apelle_acom = str_apelle_acom.Replace("Apellidos y Nombre:", "");


                    str_dni_acom = datos_acomp.Substring(pos_dni_acom, pos_fecha_nacimiento_acom - pos_dni_acom).Trim();
                    str_dni_acom = str_dni_acom.Replace("DNI:", "");
                    str_dni_acom = formatearnif(CalculaNIF(str_dni_acom));


                    str_fecha_nacimiento_acom = datos_acomp.Substring(pos_fecha_nacimiento_acom, pos_Sexo_acom - pos_fecha_nacimiento_acom).Trim();
                    str_fecha_nacimiento_acom = str_fecha_nacimiento_acom.Replace("Fecha de Nacimiento:", "");

                    str_Sexo_acom = datos_acomp.Substring(pos_Sexo_acom).Trim();
                    str_Sexo_acom = str_Sexo_acom.Replace("Sexo:", "");

                }
                else
                {

                    str_Plazas = "1";
                    str_direccion = split[i].Substring(pos_direccion).Trim();
                    dire = str_direccion.Replace("  - ", "#");

                    split_direc = dire.Split(new Char[] { '#' });

                    direc_final_dos = split_direc[0].Replace("    ", "#");

                    MessageBox.Show(direc_final_dos);

                    split_direc_uno = direc_final_dos.Split(new Char[] { '#' });



                    str_direccion_calle = split_direc_uno[0];

                    //MessageBox.Show(split_direc_uno[1]);

                    try
                    {
                        cp = split_direc_uno[1];
                    }
                    catch
                    {

                        cp = "00000";
                    }



                    str_direccion_calle = str_direccion_calle.Replace("Dirección:", "");
                    poblacion = split_direc[1];
                    //

                }

                //MessageBox.Show(ruta);

                texto_final += Environment.NewLine + expediente_final + "\t" + str_dni + "\t" + str_apelle_nombre + "\t" + str_dni_acom + "\t" + str_apelle_acom + "\t" + obtener_plazas(str_Plazas) + "\t" + fecha_inicio + "\t" + fecha_inicio + "\t" + duracion + "\t    \t     \tNO" + "\t" + str_telefono + "\t    \t     \t      \t" + poblacion + "\t" + poblacion + "\t" + cp + "\t" + str_direccion_calle;
                //MessageBox.Show(Convert.ToString(i) + "_" + (split.Length));
                if (i == (split.Length - 1))
                {


                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(ruta + @"\transformado.txt", true))
                    {

                        file.Write(texto_final);
                    }

                }

            }



            MessageBox.Show("Proceso Finalizado");
        }




        public string obtener_plazas(string cad)
        {

            string plaza = "";

            switch (cad)
            {
                case "1":
                    plaza = "A";
                    break;
                case "2":
                    plaza = "C";
                    break;
                default:
                    plaza = "C";
                    break;
            }

            return plaza;

        }


        public string leer_template(string ruta)
        {


            int counter = 0;
            string line;
            string texto = "";

            // Read the file and display it line by line.
            System.IO.StreamReader file =
             new System.IO.StreamReader(ruta + @"\template.txt");
            while ((line = file.ReadLine()) != null)
            {
                texto += line;
                counter++;
            }

            file.Close();

            return texto;



        }


        /// <summary>
        /// Dado un DNI obtiene la letra que le corresponde al NIF
        /// </summary>
        /// <param name="strA">DNI</param>
        /// <returns>Letra del NIF</returns>
        private String CalculaNIF(String strA)
        {
            const String cCADENA = "TRWAGMYFPDXBNJZSQVHLCKE";
            const String cNUMEROS = "0123456789";

            Int32 a = 0;
            Int32 b = 0;
            Int32 c = 0;
            Int32 NIF = 0;
            StringBuilder sb = new StringBuilder();

            strA = strA.Trim();
            if (strA.Length == 0) return "";

            // Dejar sólo los números
            for (int i = 0; i <= strA.Length - 1; i++)
                if (cNUMEROS.IndexOf(strA[i]) > -1) sb.Append(strA[i]);

            strA = sb.ToString();
            a = 0;
            NIF = Convert.ToInt32(strA);
            do
            {
                b = Convert.ToInt32((NIF / 24));
                c = NIF - (24 * b);
                a = a + c;
                NIF = b;
            } while (b != 0);

            b = Convert.ToInt32((a / 23));
            c = a - (23 * b);
            return strA.ToString() + cCADENA.Substring(c, 1);
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            this.textBox3.Clear();
        }





        public string obtener_balneario_admin(string id_balneario)
        {



            string admin = "";





            DataTable tabla_existe_registro = new DataTable();
            tabla_existe_registro = this.misfuncionesBDExpedientes.obtenerDatable("SELECT admin FROM [config_balneario_prereservas] WHERE id='" + id_balneario + "' ORDER BY balneario ASC");

            foreach (DataRow dr in tabla_existe_registro.Rows)
            {




                admin = Convert.ToString(dr["admin"]);






            }


            return admin;

        }



        public bool ValidateBirthday(String date)
        {
            DateTime Temp;

            if (DateTime.TryParse(date, out Temp) == true &&
          Temp.Hour == 0 &&
          Temp.Minute == 0 &&
          Temp.Second == 0 &&
          Temp.Millisecond == 0 &&
          Temp > DateTime.MinValue)
                return true;
            else
                return false;
        }



        public bool CheckBird(string nif)
        {


            this.msqlConnection = new MySql.Data.MySqlClient.MySqlConnection("server=192.168.23.233;user id=admin_bd;Password=DiodeDiode2019;database=crm;persist security info=False");
            DataTable mitable = new DataTable();

            ArrayList ArrylistContact = new ArrayList();
            string sql = "SELECT birthdate  FROM  contacts WHERE  id_nif='" + nif + "'";
            //MySqlDataAdapter mda = new MySqlDataAdapter("SELECT first_name,last_name,primary_address_street,primary_address_city,primary_address_state,primary_address_postalcode,birthdate FROM  contacts WHERE id_nif='" + relacionado_solcitud + "'", this.msqlConnection);
            MySqlDataAdapter mda = new MySqlDataAdapter(sql, this.msqlConnection);
            DataSet ds = new DataSet();
            bool birthdate = false;

            try
            {
                mda.Fill(ds);
                mitable = ds.Tables[0];



                foreach (System.Data.DataRow dr in mitable.Rows)
                {


                    try
                    {


                        if (Convert.ToDateTime(dr["birthdate"]).Year == 1970)
                        {
                            birthdate = false;

                        }
                        else
                        {

                            birthdate = true;


                        }

                    }
                    catch
                    {

                        birthdate = false;
                    }

                }
            }
            catch
            {

                birthdate = false;

            }



            return birthdate;
        }


        private void button7_Click(object sender, EventArgs e)
        {
           this.textBox1.Text= new Utils().Encrypt(ConfigurationManager.AppSettings["ConnectionString"], "balneario");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "balneario")
            {

                Encriptar.Visible = true;
            }
            else {

                Encriptar.Visible = false;
            }

        }

       
    }





}
