using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Data.SqlClient;
using System.IO;
using System.Collections;
using System.Threading;
using System.Net.Security;

using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Net.NetworkInformation;

using System.Management;
using Microsoft.Win32;
using System.Net;
using System.Net.Mail;

namespace RegistraFactura
{
    public partial class Form1 : Form
    {
        /*
          20220927 Cuando se crea el archivo de cancelacion de una factura lo crea con el mismo nombre que una cancelacion anterior, sobre escribiendo el archivo de la cancelacion anterior
         * se agrega a el nombre del archivo la factura que está siendo cancelada.
         * 
         * 20221201 Se agrega el campo esdesflote a la tabla Centralizacionv2.dbo.dig_facturaaux el campo se debe leer de SER_VEHICULO y cuando sea desflote no debe crear el archivo de facturacion, solo logueo.
         * 20230502 Solo marcamos cuando es desflote, pero si procesamos todo igual. Ver el caso de esta unidad: select * from Centralizacionv2.dbo.DIG_FACTURAAUX where vte_serie = '3N1CN7AD0JK437020' order by id_local 
         //el desflote solo fue marcado en la segunda facturación y en la tercera facturacion no se marcó como desflote.
        */

        // "C:\AndradeGPO\ActualizarCampoEnBP\Ejecutable\BusinessProSICOP.exe" SICOP GMI GAZM_Zaragoza Exporta C:\AndradeGPO\ActualizarCampoEnBP\SiCoP\Generar\ parametro_ocioso.txt 3N1CK3CD9DL259265 1000 5063

        string ConnectionString = System.Configuration.ConfigurationManager.AppSettings["ConnectionString"];
        ConexionBD objDB = null;
        string ConnectionStringTablaPaso = System.Configuration.ConfigurationManager.AppSettings["ConnectionStringTablaPaso"];
        ConexionBD objDBTablaPaso = null;

        //20150429 string RutaEjecutableBPro = System.Configuration.ConfigurationSettings.AppSettings["RutaEjecutableBPro"];
        //20150429 string DirectorioArchivosSICOP = System.Configuration.ConfigurationSettings.AppSettings["DirectorioArchivosSICOP"]; = carpeta_local_ventas
        //20150429 string Mascara = System.Configuration.ConfigurationSettings.AppSettings["Mascara"];


        string Latencia = System.Configuration.ConfigurationManager.AppSettings["Latencia"];
        string MinutosEsperaraBPro = System.Configuration.ConfigurationManager.AppSettings["MinutosEsperaraBPro"];
        string TotalIntentos = System.Configuration.ConfigurationManager.AppSettings["TotalIntentos"];
        string TopeMinBusqenBDSicop = System.Configuration.ConfigurationManager.AppSettings["TopeMinBusqenBDSicop"];

        string NumeroSucursalProcesar = System.Configuration.ConfigurationManager.AppSettings["NumeroSucursalProcesar"];    
        //ArrayList fw = new ArrayList();

        Dictionary<string, Thread> dicHilos = new Dictionary<string, Thread>();

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

        public Form1()
        {
            InitializeComponent();
        }


        private void ProcesaBitacora()
        {
            string Q = "Select * from DIG_FACTURAAUX where id_estatus=1";
            if (this.NumeroSucursalProcesar.Trim() != "")
                Q += " and suc_idsucursal=" + this.NumeroSucursalProcesar.Trim();
            //Q += " and id_maquina=32";            
            Q += " order by id_local";

            DataSet ds = this.objDBTablaPaso.Consulta(Q);
            if (!this.objDBTablaPaso.EstaVacio(ds))
            {
                foreach (DataRow lector in ds.Tables[0].Rows)
                {
                    try
                    {
                        string suc_idsucursal = lector["suc_idsucursal"].ToString().Trim();
                        string vte_docto = lector["vte_docto"].ToString().Trim();
                        string vte_fechdocto = lector["vte_fechdocto"].ToString().Trim();
                        string vte_serie = lector["vte_serie"].ToString().Trim();
                        string aux_fecha = lector["aux_fecha"].ToString().Trim();
                        string proceso = lector["proceso"].ToString().Trim();
                        string id_local = lector["id_local"].ToString().Trim();
                        string canal_venta = lector["canal_venta"].ToString().Trim();

                        const string quote = "\"";

                        //consultamos el id_agencia de la tabla de equivalencias.
                        Q = "Select id_agencia from SICOP_AGENCIA_SUCURSAL where suc_idsucursal=" + suc_idsucursal.Trim();
                        string id_agencia = this.objDB.ConsultaUnSoloCampo(Q);

                        if (id_agencia.Trim() != "")
                        {
                            #region Cancelacion
                            if (proceso.Trim() == "cancelacion")
                            {
                                //Si la primer letra es A es factura de Auto Nuevo 
                                string tipo_auto = vte_docto.Substring(0, 1) == "A" ? "NUEVO" : "SEMINUEVO";
                                string saux_tipo = tipo_auto == "NUEVO" ? "N'" : "S%'";
                                Q = "Select dealerid from SICOP_DEALER_AGENCIA where id_agencia=" + id_agencia + " and tipo like '" + saux_tipo.Trim();
                                string id_distribuidor = this.objDB.ConsultaUnSoloCampo(Q);
                                if (id_distribuidor.Trim() != "")
                                {
                                    string LineaEncabezado = quote + "IdDistribuidor" + quote + "," + quote + "Factura" + quote + "," + quote + "Vin" + quote;
                                    string LineaDetalle = quote + id_distribuidor.Trim() + quote + "," + quote + vte_docto.Trim() + quote + "," + quote + vte_serie.Trim() + quote;
                                    string CarpetaRemota = this.objDB.ConsultaUnSoloCampo("Select carpeta_remota From SICOPCONFIGXMAQUINA where activo='True' and numero_sucursal='" + id_agencia.Trim() + "'").Trim();
                                    if (tipo_auto == "SEMINUEVO")
                                          CarpetaRemota = CarpetaRemota + "\\SEMINUEVOS";
                                    
                                    //20220927
                                    //string rutaArchivo = CarpetaRemota + "\\Cancelacion\\SICOP_CANCELACION_TEMP_DMS_" + vte_serie.Trim() + ".TXT";
                                    string rutaArchivo = CarpetaRemota + "\\Cancelacion\\SICOP_CANCELACION_TEMP_DMS_" + vte_serie.Trim() + "_" + vte_docto.Trim() + ".TXT";
                                    if (!Directory.Exists(CarpetaRemota +"\\Cancelacion"))
                                        Directory.CreateDirectory(CarpetaRemota +"\\Cancelacion"); 
                                    
                                    CreaArchivoCancelacion(LineaEncabezado, LineaDetalle, rutaArchivo);
                                    if (File.Exists(rutaArchivo))
                                    {
                                        Q = "Update DIG_FACTURAAUX set id_estatus=2, fh_estatus=getdate(), fh_archseekop=getdate(), rutaarchseekop='" + rutaArchivo .Trim() + "'  where id_local=" + id_local.Trim();
                                        this.objDBTablaPaso.EjecUnaInstruccion(Q);
                                    }
                                    else { //Estatus 3 quiere decir que hubo un error en la creacion del archivo.
                                        Q = "Update DIG_FACTURAAUX set id_estatus=3, fh_estatus=getdate() where id_local=" + id_local.Trim();
                                        this.objDBTablaPaso.EjecUnaInstruccion(Q);                                                                            
                                    }
                                }
                            }//del proceso de Cancelacion.
                            #endregion
                            #region VEnta
                            if (proceso.Trim() == "venta")
                            {
                                //20221211 Primero averiguamos si se trata de un desflote
                                string EsDesflote = "0"; //si es null o 0 el campo en ser_vehiculo NO ES DESFLOTE.
                                Q = "select nombre_base from Centralizacionv2.dbo.DIG_CAT_BASES_BPRO where suc_idsucursal = " + suc_idsucursal.Trim();
                                string NombreBaseBP = this.objDBTablaPaso.ConsultaUnSoloCampo(Q);
                                if (NombreBaseBP.Trim() != "")
                                {
                                    Q = "select Isnull(VEH_ESDESFLOTE,0) from " + NombreBaseBP.Trim() + ".dbo.SER_VEHICULO where VEH_NUMSERIE = '" + vte_serie.Trim() + "'";
                                    EsDesflote = this.objDBTablaPaso.ConsultaUnSoloCampo(Q);
                                }

                                //20230502 Solo marcamos cuando es desflote, pero si procesamos todo igual. Ver el caso de esta unidad: select * from Centralizacionv2.dbo.DIG_FACTURAAUX where vte_serie = '3N1CN7AD0JK437020' order by id_local 
                                //el desflote solo fue marcado en la segunda facturación y en la tercera facturacion no se marcó como desflote.

                                //if (EsDesflote.Trim() == "0" || EsDesflote.Trim() == "")
                                //{//solo debe crear archivo de facturacion cuando no es desflote:

                                    Q = "Select ruta_ejecutable_BPro From SICOPCONFIGXMAQUINA where activo='True' and numero_sucursal='" + id_agencia.Trim() + "'";
                                    string RutaEjecutableBPro = this.objDB.ConsultaUnSoloCampo(Q).Trim();
                                    if (RutaEjecutableBPro.Trim() != "")
                                    {
                                        string DirectorioArchivosSICOP = this.objDB.ConsultaUnSoloCampo("Select carpeta_local_ventas From SICOPCONFIGXMAQUINA where activo='True' and numero_sucursal='" + id_agencia.Trim() + "'").Trim();
                                        string Sicop = "SICOP";
                                        string Comando = quote + RutaEjecutableBPro.Trim() + quote + " {0} {1} {2} {3} {4} {5} {6}"; //this.RutaEjecutableBPro.Trim() + " {0} {1} {2} {3} {4} {5} {6}";   //@""" + this.RutaEjecutableBPro.Trim() + @""" + " {0} {1} {2} {3} {4} {5} {6}";
                                        string Sentido = "Exporta";
                                        string UsuarioBPRo1 = this.objDB.ConsultaUnSoloCampo("Select usuario_bpro From SICOPCONFIGXMAQUINA where activo='True' and numero_sucursal='" + id_agencia.Trim() + "'").Trim();
                                        string BDBPRo1 = this.objDB.ConsultaUnSoloCampo("Select bd_bpro From SICOPCONFIGXMAQUINA where activo='True' and numero_sucursal='" + id_agencia.Trim() + "'").Trim();
                                        string carpeta_local_ventas = this.objDB.ConsultaUnSoloCampo("Select carpeta_local_ventas From SICOPCONFIGXMAQUINA where activo='True' and numero_sucursal='" + id_agencia.Trim() + "'").Trim();
                                        string mascara = this.objDB.ConsultaUnSoloCampo("Select mascara From SICOPCONFIGXMAQUINA where activo='True' and numero_sucursal='" + id_agencia.Trim() + "'").Trim();

                                        try
                                        {
                                            //"C:\Users\omorales\Desktop\Business Pro SICOP.exe" SICOP GMI GAZM_ZARAGOZA Exporta C:\SiCoP\Generar\ SICOP_PROSPECTOS_TEMP_DMS.TXT 3N1CK3CD9DL259265 1000 25832
                                            //"C:\Users\omorales\Desktop\Business Pro SICOP.exe" SICOP GMI GAZM_ZARAGOZA Exporta C:\SiCoP\Generar\ SICOP_PROSPECTOS_TEMP_DMS.TXT 3N1CK3CD9DL259265
                                            Comando = string.Format(Comando, Sicop, UsuarioBPRo1, BDBPRo1, Sentido, DirectorioArchivosSICOP.Trim(), "parametro_ocioso.txt", vte_serie.Trim());
                                            LanzaEjecucion(Comando); //lo deja en una sola carpeta.                                 
                                            Utilerias.WriteToLog("Se ejecutó: " + Comando, "ProcesaBitacora", Application.StartupPath + "\\Log.txt");
                                            //Esperamos un minuto para que le de tiempo a la interfaz a crear el archivo.
                                            Thread.Sleep(Convert.ToInt16(this.MinutosEsperaraBPro) * 60000); //20150505 En lugar del fsw_created.
                                            string ArchivoSeekop = procesaArchivoGeneradoporBPro(carpeta_local_ventas.Trim() + mascara.Trim(), mascara.Trim(), id_local.Trim(), id_agencia.Trim());
                                            if (ArchivoSeekop.Trim() != "")
                                            {
                                                if (File.Exists(ArchivoSeekop))
                                                {
                                                    Q = "Update DIG_FACTURAAUX set id_estatus=2, fh_estatus=getdate(), fh_archseekop=getdate(), rutaarchseekop='" + ArchivoSeekop.Trim() + "', VEH_ESDESFLOTE = '" + EsDesflote.Trim() + "'  where id_local=" + id_local.Trim();
                                                    this.objDBTablaPaso.EjecUnaInstruccion(Q);
                                                }
                                                else
                                                { //Estatus 3 quiere decir que hubo un error en la creacion del archivo.
                                                    Q = "Update DIG_FACTURAAUX set id_estatus=3, fh_estatus=getdate(), VEH_ESDESFLOTE = '" + EsDesflote.Trim() + "' where id_local=" + id_local.Trim();
                                                    this.objDBTablaPaso.EjecUnaInstruccion(Q);
                                                }
                                            }
                                            else
                                            {
                                                //Estatus 3 quiere decir que hubo un error en la creacion del archivo.
                                                Q = "Update DIG_FACTURAAUX set id_estatus=3, fh_estatus=getdate(), VEH_ESDESFLOTE = '" + EsDesflote.Trim() + "' where id_local=" + id_local.Trim();
                                                this.objDBTablaPaso.EjecUnaInstruccion(Q);
                                                Utilerias.WriteToLog("Validar generacion del Archivo con Factura: " + vte_docto.Trim() + " vin: " + vte_serie + " id_local: " + id_local + " id_agencia: " + id_agencia + " canal_venta: " + canal_venta, "ProcesaBitacora", Application.StartupPath + "\\Log.txt");
                                            }
                                            Utilerias.WriteToLog("", "", Application.StartupPath + "\\Log.txt");
                                            Thread.Sleep(20000);
                                        }
                                        catch (Exception ex)
                                        {
                                            Debug.WriteLine(ex.Message);
                                            Utilerias.WriteToLog(ex.Message, "ProcesaBitacora", Application.StartupPath + "\\Log.txt");
                                        }
                                    }//de si existe ruta del ejecutable de BPro.
                                    else
                                    {
                                        Utilerias.WriteToLog(" No se encontró ruta de ejecutable de BPro para la agencia: " + id_agencia, "ProcesaBitacora", Application.StartupPath + "\\Log.txt");
                                    }
                                
                                //}//De si es desflote
                                //else {
                                //    Q = "Update DIG_FACTURAAUX set id_estatus=4, fh_estatus=getdate(),VEH_ESDESFLOTE = '" + EsDesflote.Trim() + "' where id_local=" + id_local.Trim();
                                //    this.objDBTablaPaso.EjecUnaInstruccion(Q);
                                //    Utilerias.WriteToLog("ES DESFLOTE: No se creá el archivo de Facturacion en la carpeta SEEKOP Factura: " + vte_docto.Trim() + " vin: " + vte_serie + " id_local: " + id_local + " id_agencia: " + id_agencia, "ProcesaBitacora", Application.StartupPath + "\\Log.txt");
                                //}

                            } //del proceso de venta                    
                            #endregion
                        }
                        else
                        {
                            Utilerias.WriteToLog("No se encontro id_agencia para la suc_idsucursal " + suc_idsucursal.Trim() + " en SICOP_AGENCIA_SUCURSAL", "ProcesaBitacora", Application.StartupPath + "\\Log.txt");
                        }

                    }//del try
                    catch (Exception ex)
                    {
                        Utilerias.WriteToLog(ex.Message, "ProcesaBitacora", Application.StartupPath + "\\Log.txt");
                        Debug.WriteLine(ex.Message);
                    }
                } //del ciclo sobre cada registro.
            } //de que hay registros por procesar.
        } // de Procesa Bitacora

        private void Form1_Load(object sender, EventArgs e)
        {            
            FileInfo archivoExecutalbe = new FileInfo(Application.ExecutablePath.Trim());
            string NombreProceso = archivoExecutalbe.Name; 
            NombreProceso = NombreProceso.Replace(".exe", "");
            NombreProceso = NombreProceso.Replace(".EXE", "");
            if (CuentaInstancias(NombreProceso) == 1)
            {//la instancia debe ser igual a 1, que es esta misma instancia. Si es distinta entonces mandar el aviso de que ya se está ejecutando.
                Utilerias.WriteToLog("", "", Application.StartupPath + "\\Log.txt");  
                
                this.objDB = new ConexionBD(this.ConnectionString.Trim());
                this.objDBTablaPaso = new ConexionBD(this.ConnectionStringTablaPaso.Trim()); 

                this.timerThread.Interval=(Convert.ToInt16(this.Latencia) * 60000);
                this.timerThread.Enabled = true;
                this.timerThread.Start();
                
                this.timerReproceso.Interval = 3600000; //1 hora
                this.timerReproceso.Enabled = true; 
                this.timerReproceso.Start(); 
                //ProcesaBitacora();
                /*
                    string Q = "Select numero_sucursal, mascara, carpeta_local_ventas, ruta_ejecutable_BPro ";
                    Q += " From SICOPCONFIGXMAQUINA ";
                    Q += " where activo='True' order by Convert(int,numero_sucursal)"; 
                    
                    DataSet ds = this.objDB.Consulta(Q);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {                        

                        foreach (DataRow reg in ds.Tables[0].Rows)
                        {
                            //20150429 this.fsw.Path = this.DirectorioArchivosSICOP.Trim(); //TODO: por parametrizar en BD. i.e. crear uno por agencia activa.
                            //20150429 this.fsw.Filter = this.Mascara.Trim();
                            //http://bytes.com/topic/c-sharp/answers/227562-multiple-filesystemwatchers
                            // otra posible solución: http://www.codeproject.com/Articles/271669/Using-FileSystemWatcher-to-monitor-multiple-direct
                            try
                            {                                                                
                                FileSystemWatcher Clientwatcher = new FileSystemWatcher();
                                Clientwatcher.Path = reg["carpeta_local_ventas"].ToString().Trim();
                                Clientwatcher.Filter = reg["mascara"].ToString().Trim();
                                Clientwatcher.NotifyFilter = NotifyFilters.FileName;
                                Clientwatcher.Created += new FileSystemEventHandler(this.fsw_Created);
                                //Clientwatcher.Changed += new FileSystemEventHandler( ClientFileUpdated);
                                Clientwatcher.EnableRaisingEvents = true;
                                this.components.Add(Clientwatcher);
                                Utilerias.WriteToLog("Se estableció el visor de la carpeta: " + reg["carpeta_local_ventas"].ToString().Trim() + " con máscara: " + reg["mascara"].ToString().Trim(), "Form1_Load", Application.StartupPath + "\\Log.txt");                                 
                            }
                            catch (Exception ex)
                            {
                                Utilerias.WriteToLog("Error al crear el fsw para la agencia: " + reg["numero_sucursal"].ToString().Trim() + ex.Message, "Form1_Load", Application.StartupPath + "\\Log.txt");
                            }
                        }
                    } 
                Utilerias.WriteToLog("", "", Application.StartupPath + "\\Log.txt"); 
                */

                //RevisaLectores();
                //ProcesaBitacora();
            }
            else
            {
                //Utilerias.WriteToLog("Ya existe una instancia de: " + NombreProceso + " se conserva la instancia actual", "Form1_Load", Application.StartupPath + "\\Log.txt");
                Application.Exit();
            }
        }
    
    

        /// <summary>
        /// En la base de datos de BPro hace update al campo fechaentregareal
        /// </summary>
        /// <param name="CodigoLeido"></param>
        /// <returns>Si fue exito o error. Error cuando ya con anterioridad se registro la fecha de salida.</returns>
        public string RegistrarSalida(string vin,string fecharegistrar, string id_agencia, string id_bitacora)
        {
            //3N6DD21T0DK077800
            //1
                       
            string res = "";
            string Q = "";

            try
            {
                //primero analizamos la cadena capturada y si tiene el formato requerido la parseamos.
                
                if (vin.Length > 0)
                {                                        
                    //Consultamos los datos para poder firmarnos en la base de datos.
                    #region Consulta de los datos para el Logueo en el Servidor Remoto 
                    //ConexionBDchkDos objDB = new ConexionBDchkDos(this.CadenaConexion);
                    //SqlConnection conBP = new SqlConnection();
                    //SqlCommand bp_comand = new SqlCommand();
                    

                    //conociendo el id_agencia procedemos a consultar los datos de conexion en la tabla transferencia
                    Q = "Select ip,usr_bd,pass_bd,nombre_bd,bd_alterna, dir_remoto_xml, dir_remoto_pdf,usr_remoto,pass_remoto, ip_almacen_archivos, smtpserverhost, smtpport, usrcredential, usrpassword, actualizafechappago ";
                    Q += " From SICOP_TRASMISION where id_agencia='" + id_agencia + "'";

                    DataSet ds = this.objDB.Consulta(Q);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        DataRow regConexion = ds.Tables[0].Rows[0];
                        string strconexionABussinesPro = string.Format("Data Source={0};Initial Catalog={1}; Persist Security Info=True; User ID={2};Password={3}", regConexion["ip"].ToString(), regConexion["nombre_bd"].ToString(), regConexion["usr_bd"].ToString(), regConexion["pass_bd"].ToString());                        
                        string ActualizaFechaPromesaPago = regConexion["actualizafechappago"].ToString().ToUpper().Trim();
                        string nombre_bd=regConexion["nombre_bd"].ToString();

                        //Utilerias.WriteToLog(" ActualizaFechaPromesaPago :" + ActualizaFechaPromesaPago, "RegistrarSalida", Application.StartupPath + "\\Log.txt"); 

                        //if (conBP.State.ToString().ToUpper().Trim() == "CLOSED")
                        //{
                            try
                            {
                                //conBP.ConnectionString = strconexionABussinesPro;
                                //conBP.Open();
                                //bp_comand.Connection = conBP;

                                ConexionBD objDBBPro = new ConexionBD(strconexionABussinesPro); 

                                //teniendo la conexion con la base de datos vamos a actualizar
                                //20210930 Se consulta la factura desde BPRO y no desde el archivo de facturacion, porque el mismo no se genera para Flotilla e Intercambios
                                string factura = "";
                                string fecha_factura = "";
                                Q = " select VTE_DOCTO, Convert(char(8),Convert(date,VTE_FECHDOCTO),112) as VTE_FECHDOCTO from ADE_VTAFI";
                                Q += " WHERE VTE_STATUS='I' AND VTE_SERIE='" + vin.Trim() + "'";
                                //string factura = objDBBPro.ConsultaUnSoloCampo(Q).Trim(); 
                                DataSet dsFac = objDBBPro.Consulta(Q);
                                if (!objDBBPro.EstaVacio(dsFac))
                                {
                                    DataRow reg = dsFac.Tables[0].Rows[0];
                                     factura = reg["VTE_DOCTO"].ToString().Trim();
                                     fecha_factura = reg["VTE_FECHDOCTO"].ToString().Trim();

                                    if (factura.Trim() != "")
                                    {
                                        Q = "Update SICOP_BITACORA set factura = '" + factura.Trim() + "', fecha_factura='" + fecha_factura.Trim()  + "' where id_bitacora=" + id_bitacora.Trim() + " and aquien='" + vin.Trim() + "'";
                                        this.objDB.EjecUnaInstruccion(Q);
                                    }
                                }
                                //20200722 es necesario identificar primero si es unidad nueva o seminueva.
                                // si la unidad es seminueva la actualizacion debe ir sobre  USN_PEDIDO..PMS_FECHAREALENTREGA
                                string Seminuevo = "";
                                Q = "select VEH_SITUACION from SER_VEHICULO  where VEH_NUMSERIE = '" + vin.Trim() + "'";
                                string Situacion = objDBBPro.ConsultaUnSoloCampo(Q);
                                if (Situacion.Trim() != "")
                                {
                                    Situacion = Situacion.Substring(0, 1);
                                    Situacion = Situacion.ToUpper();
                                    Seminuevo = Situacion.Trim() == "S" ? "Seminuevo" : "Nuevo";
                                    Q = "Update SICOP_BITACORA set tipo_auto = '" + Seminuevo.Trim() + "' where id_bitacora=" + id_bitacora.Trim() + " and aquien='" + vin.Trim() + "'";
                                    this.objDB.EjecUnaInstruccion(Q);
                                }
                                
                                string id_pedido = "";
                                string QConsulta = "";

                                if (Seminuevo == "Seminuevo")
                                {
                                    //buscamos el pedido de la unidad 
                                    Q = "Select PMS_NUMPEDIDO from USN_PEDIDO where PMS_NUMSERIE ='" + vin.Trim() + "'";
                                    Q += " and PMS_STATUS='I'";
                                    id_pedido = objDBBPro.ConsultaUnSoloCampo(Q);

                                    Q = " update USN_PEDIDO set PMS_FECHAREALENTREGA = '" + fecharegistrar + "'";
                                    Q += " where PMS_NUMSERIE ='" + vin.Trim() + "'";
                                    Q += " and Isnull(PMS_FECHAREALENTREGA,'01/01/1900')='01/01/1900'"; //para que solo lo haga una sola vez.
                                    Q += " and PMS_NUMPEDIDO = '" + id_pedido.Trim() + "'";  //no todos solo el pedido vivo de la unidad. 20200724

                                    QConsulta = " select Isnull(PMS_FECHAREALENTREGA,'') from USN_PEDIDO where PMS_NUMSERIE ='" + vin.Trim() + "'";
                                    QConsulta += " and PMS_NUMPEDIDO = '" + id_pedido.Trim() + "'";

                                }
                                else
                                {   //Nuevos
                                    //buscamos el pedido de la unidad 20200724
                                    Q = "Select PEN_IDPEDI from UNI_PEDIUNI where pen_numserie ='" + vin.Trim() + "'";
                                    Q += " and PEN_STATUS='I'";
                                    id_pedido = objDBBPro.ConsultaUnSoloCampo(Q);                                      
                                    
                                    //Q = "update UNI_PEDIUNI set PEN_FECHAENTREGA = Convert(char(10),getdate(),103) where pen_idpedi ='" + id_pedido.Trim()  + "' and pen_numserie ='" + vin.Trim() +  "'";
                                    //Q = "update UNI_PEDIUNI set PEN_FECHAENTREGA_REAL = Convert(char(10),getdate(),103) where pen_idpedi ='" + id_pedido.Trim() + "' and pen_numserie ='" + vin.Trim() + "'";
                                    //Q = "update UNI_PEDIUNI set PEN_FECHAENTREGA_REAL = Convert(char(10),getdate(),103) ";
                                    Q = " update UNI_PEDIUNI set PEN_FECHAENTREGA_REAL = '" + fecharegistrar + "'";
                                    Q += " where pen_numserie ='" + vin.Trim() + "'";
                                    Q += " and Isnull(PEN_FECHAENTREGA_REAL,'01/01/1900')='01/01/1900'";//para que solo lo haga una sola vez.
                                    Q += " and pen_idpedi ='" + id_pedido.Trim() + "'"; //no todos solo el pedido vivo de la unidad. 20200724
                                    
                                    QConsulta = "Select Isnull(PEN_FECHAENTREGA_REAL,'') from UNI_PEDIUNI where pen_numserie ='" + vin.Trim() + "'";
                                    QConsulta += " and pen_idpedi ='" + id_pedido.Trim() + "'";
                                }
                                
                                int totreg = objDBBPro.EjecUnaInstruccion(Q);
                                Utilerias.WriteToLog("Se seteo la fecha real de entrega: " + Q.Trim(), "RegistrarSalida", Application.StartupPath + "\\Log.txt");                                 
                                //---
                                    string FechaSeteada = "";
                                    int Contador = 1;
                                  while (FechaSeteada.Trim() == "" && Contador < 6)
                                   {
                                     FechaSeteada = objDBBPro.ConsultaUnSoloCampo(QConsulta).Trim();
                                     if (FechaSeteada.Trim() == "")
                                     {
                                         totreg = objDBBPro.EjecUnaInstruccion(Q);
                                         Utilerias.WriteToLog("Se seteo la fecha real de entrega: " + Q.Trim(), "RegistrarSalida", Application.StartupPath + "\\Log.txt");
                                         Thread.Sleep(20000);
                                     }
                                     else {
                                         Utilerias.WriteToLog("Intento : " + Contador.ToString()  + " de seteo de fecha real de entrega: " + Q.Trim(), "RegistrarSalida", Application.StartupPath + "\\Log.txt");
                                     }
                                     if (Contador == 5)
                                     {
                                         EnviaCorreoSISTEMAS(vin.Trim(), id_agencia.ToString(), id_bitacora.ToString(), Seminuevo, nombre_bd.Trim(), Q.Trim()); 
                                     }

                                   Contador++;         
                                   }
                                
                                //bp_comand.CommandText = Q.Trim();                                
                                //int totreg = bp_comand.ExecuteNonQuery();                                

                                  if (FechaSeteada.Trim() != "")
                                {
                                    res = "La salida de la unidad con vin: " + vin.Trim() + " ha quedado registrada.";
                                }
                                else {
                                    res = "La fecha de salida ya fue capturada: " + vin.Trim();
                                }

                                  string TipoVenta = "";

                                  if (Seminuevo == "Nuevo")
                                  {
                                      //20170202 Consultamos el tipo de venta de la unidad.
                                      //Q = " select VTE_DOCTO,VTE_STATUS,VTE_FORMAPAGO,PAR_DESCRIP1 from ADE_VTAFI,PNC_PARAMETR";
                                      Q = " select PAR_DESCRIP1 from ADE_VTAFI,PNC_PARAMETR";
                                      Q += " WHERE VTE_STATUS='I' AND PAR_TIPOPARA='VNT' AND VTE_FORMAPAGO=PAR_IDENPARA AND VTE_SERIE='" + vin.Trim() + "'";                                      
                                      TipoVenta = objDBBPro.ConsultaUnSoloCampo(Q);
                                  }
                                  if (Seminuevo == "Seminuevo")
                                  { 
                                      //20210215 Consultamos el canal de venta de seminuevos.                                      
                                      Q = " select PAR_DESCRIP1 from PNC_PARAMETR, USN_PEDIDO where ";
                                      Q += " PAR_tipoPARA = 'TIPVENUSA' AND PAR_IDENPARA=PMS_TIPOVENTA"; 
                                      Q += " AND PMS_NUMSERIE = '" + vin.Trim() +"'";
                                      Q += " AND PMS_NUMPEDIDO = " + id_pedido.Trim();
                                      TipoVenta = objDBBPro.ConsultaUnSoloCampo(Q);
                                  }
                                  if (TipoVenta.Trim() != "")
                                  {
                                      Q = "Update SICOP_BITACORA set tipo_venta = '" + TipoVenta.Trim() + "' where id_bitacora=" + id_bitacora.Trim() + " and aquien='" + vin.Trim() + "'";
                                      this.objDB.EjecUnaInstruccion(Q);
                                  }


                                
                                //20180828 Centralizacion requiere colocar la fecha promesa de pago en la cartera.
                                #region Colocando Fecha Promesa de Pago. 20180829
                                try
                                {

                                    if (ActualizaFechaPromesaPago == "TRUE")
                                    {
                                        //SqlConnection conBPConcentra = new SqlConnection();
                                        //SqlCommand bp_comandConcentra = new SqlCommand();
                                        //Primero validamos que no se haya capturado antes la fecha promesa de pago para este vin.
                                        Q = " Select Convert(char(8),isnull(fecha_promesa_pago,'19000101'),112) from SICOP_BITACORA where aquien='" + vin.Trim() + "'";
                                        string fechadesdeBD = this.objDB.ConsultaUnSoloCampo(Q).Trim();
                                        if (fechadesdeBD.Trim() == "19000101")
                                        {
                                            string strTipoPolizacsv = "";

                                            Q = "select PAR_IDENPARA from  [PNC_PARAMETR] where PAR_TIPOPARA='TIPOLI'";
                                            Q += "and PAR_DESCRIP2='CU' and PAR_IDMODULO='CON' and PAR_STATUS='A'";
                                            DataSet dsP = objDBBPro.Consulta(Q);
                                            if (dsP != null && dsP.Tables.Count > 0 && dsP.Tables[0].Rows.Count > 0)
                                            {
                                                foreach (DataRow regP in dsP.Tables[0].Rows)
                                                {
                                                    strTipoPolizacsv += "'" + regP["PAR_IDENPARA"].ToString().Trim() + "',";
                                                }

                                                strTipoPolizacsv = strTipoPolizacsv == "" ? strTipoPolizacsv.Trim() : strTipoPolizacsv.Substring(0, strTipoPolizacsv.LastIndexOf(','));

                                                //Utilerias.WriteToLog(strTipoPolizacsv, "RegsitrarSalida_2", Application.StartupPath + "\\Log.txt");
                                            }

                                            /*
                                            string BDNomConcentradora = "";
                                            Q = " Select PAR_DESCRIP1 from  [PNC_PARAMETR] where PAR_TIPOPARA = 'CONCENTRA' and PAR_STATUS='A'";
                                            BDNomConcentradora = objDBBPro.ConsultaUnSoloCampo(Q).Trim();
                                            if (BDNomConcentradora.Trim() != "")
                                            { //Tiene concentradora La cartera está en la base concentradora.
                                                string strconexionABussinesProConcentradora = string.Format("Data Source={0};Initial Catalog={1}; Persist Security Info=True; User ID={2};Password={3}", regConexion["ip"].ToString(), BDNomConcentradora, regConexion["usr_bd"].ToString(), regConexion["pass_bd"].ToString());
                                                conBPConcentra.ConnectionString = strconexionABussinesPro;
                                                conBPConcentra.Open();
                                                bp_comandConcentra.Connection = conBPConcentra;
                                            */

                                            Q = " select top 1 CCP_FECHPROMPAG,CCP_IDPERSONA,CCP_CONSCARTERA, CCP_CONSMOV , Vcc_Anno ";
                                            Q += " from [VIS_CONCAR01] ";
                                            Q += " where CCP_OBSGEN = '" + vin.Trim() + "'";
                                            Q += " and CCP_TIPODOCTO = 'FAC'";
                                            Q += " and CCP_ABONO <> 0 ";
                                            if (strTipoPolizacsv.Trim() != "")
                                            {
                                                Q += " and CCP_TIPOPOL in (";
                                                //Q += " select PAR_IDENPARA from  [192.168.20.31].[GAZM_FAerea].[dbo].[PNC_PARAMETR] where PAR_TIPOPARA='TIPOLI'";
                                                //Q += " and PAR_DESCRIP2='CU' and PAR_IDMODULO='CON' and PAR_STATUS='A'";
                                                Q += strTipoPolizacsv.Trim();
                                                Q += ")"; //--CAB --CAF  parametro TIPOLI
                                            }
                                            Q += " order by CCP_CONSMOV desc"; //--Encontrar el maximo. top 1 desc.

                                            //Utilerias.WriteToLog(Q, "RegsitrarSalida_1", Application.StartupPath + "\\Log.txt");  

                                            DataSet dsA = objDBBPro.Consulta(Q);
                                            if (dsA != null && dsA.Tables.Count > 0 && dsA.Tables[0].Rows.Count > 0)
                                            {
                                                foreach (DataRow regA in dsA.Tables[0].Rows)
                                                {
                                                    string Vcc_Anno = regA["Vcc_Anno"].ToString().Trim();
                                                    string CCP_FECHPROMPAG = regA["CCP_FECHPROMPAG"].ToString().Trim();
                                                    string CCP_CONSCARTERA = regA["CCP_CONSCARTERA"].ToString().Trim();
                                                    string CCP_IDPERSONA = regA["CCP_IDPERSONA"].ToString().Trim();

                                                    //Utilerias.WriteToLog("Se pondra la fecha en la cartera: " + CCP_CONSCARTERA + " CCP_FECHPROMPAG: " + CCP_FECHPROMPAG, "RegistraSalida", Application.StartupPath + "\\Log.txt");  

                                                    if (CCP_CONSCARTERA != "")
                                                    { //Actualizamos la fecha promesa de pago al dia de hoy más 1.
                                                        Q = "Select CONVERT(VARCHAR(10), getdate() + 1, 103) ";
                                                        string FechaPromesaPago = this.objDB.ConsultaUnSoloCampo(Q).Trim();

                                                        //VIS_CONCAR01
                                                        Q = " Update [Con_Car01" + Vcc_Anno.Trim() + "] SET CCP_FECHPROMPAG = '" + FechaPromesaPago + "'";   //CONVERT(VARCHAR(10), getdate() + 1, 103)";
                                                        Q += " where CCP_OBSGEN = '" + vin.Trim() + "'";
                                                        Q += " and CCP_TIPODOCTO = 'FAC'";
                                                        Q += " and CCP_IDPERSONA = " + CCP_IDPERSONA.Trim();
                                                        Q += " and CCP_CONSCARTERA = " + CCP_CONSCARTERA;
                                                        int regafectados = objDBBPro.EjecUnaInstruccion(Q);
                                                        if (regafectados > 0)
                                                        {
                                                            Q = " Update SICOP_BITACORA set fecha_promesa_pago=Convert(char(8), Convert(datetime,'" + FechaPromesaPago.Trim() + "'),112) where id_bitacora=" + id_bitacora.ToString().Trim();
                                                            this.objDB.EjecUnaInstruccion(Q);
                                                            Utilerias.WriteToLog("Se actualizó la FechaPromesaDePago: " + FechaPromesaPago.Trim() + " Tenia esta fecha: " + CCP_FECHPROMPAG.Trim() + "  CCP_CONSCARTERA= " + CCP_CONSCARTERA + " vin: " + vin, "RegistrarSalida", Application.StartupPath + "\\Log.txt");
                                                        }
                                                    }//De que conocemos el consecutivo de la cartera
                                                } // de cada registro en la cartera a ponerle la fecha de pago
                                            }//de que tiene registros la cartera para el vin escaneado
                                        }//de que no se ha puesto con anterioridad la fecha promesa de pago solo debe ponerla en el primer escaneo. 20190705 
                                    } //de actualizar la fecha promesa de pago.
                                }
                                catch (Exception exC)
                                {
                                    Utilerias.WriteToLog(exC.Message +  "  Error al intentar colocar la fecha promesa de pago: ", "RegistrarSalida", Application.StartupPath + "\\Log.txt");                                 
                                }
                                #endregion
                            }
                            catch (Exception ex1)
                            {
                                res = "Error: Imposible conexion con BD de BP conexion " + strconexionABussinesPro.Trim() + "    "  + ex1.Message;
                            }
                        //} del if de la conexion.
                    }
                    else
                    {
                        res = "Error: No fue posible autenticarse en el servidor remoto";
                    }
                    #endregion                   
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                Utilerias.WriteToLog(ex.Message, "RegistraSalida", Application.StartupPath + "\\Log.txt");
            }

            Utilerias.WriteToLog(res.Trim(), "RegistraSalida", Application.StartupPath + "\\Log.txt");
            return res;
        }       


        /// <summary>
        /// Dado el vin, consulta el id_prospecto desde BPRo
        /// </summary>
        /// <param name="CodigoLeido">vin</param>
        /// <returns>vacio si no encontró al id_prospecto</returns>
        public string BuscaIDProspecto(string CodigoLeido, string NumeroSucursal)
        {
            string res = "";
            string Q = "";

            Q = "Select centralizada,nombre_bd from SICOP_TRASMISION  where centralizada = 1 and id_agencia='" + NumeroSucursal + "'";
            ConexionBD objDB59 = new ConexionBD(this.ConnectionString);            
            DataSet ds2 = objDB59.Consulta(Q);             
            
            if (!objDB59.EstaVacio(ds2) &&  ds2.Tables[0].Rows[0]["centralizada"].ToString() == "1")   
            { //Es centralizada 
                string nombre_bd = ds2.Tables[0].Rows[0]["nombre_bd"].ToString().Trim();
                Q = "Select top 1 cs.scp_idsicop from [192.168.20.29].[" + nombre_bd.Trim() + "].[dbo].[ADE_VTAFI] VT,";  
                Q += " [192.168.20.29].[GA_Corporativa].[dbo].[cat_idsicop] cs,";
                Q += " SICOP_AGENCIA_SUCURSAL ags";
                Q +=" where VT.VTE_SERIE = '" + CodigoLeido.Trim() + "'";  
                Q += " and VT.VTE_STATUS='I'";
                Q += " and VTE_TIPODOCTO = 'A' ";
                Q += " and VT.VTE_IDCLIENTE = cs.scp_idbpro";
                Q += " and cs.scp_sucursal = ags.suc_idsucursal";
                Q += " and ags.id_agencia = '" + NumeroSucursal.Trim() + "'";  
                Q += " order by cs.scp_fechope desc";
                res = objDB59.ConsultaUnSoloCampo(Q);
                  
            }
            else {
                //no es centralizada
                #region Sucursales no centralizadas
                try
                {
                    //primero analizamos la cadena capturada y si tiene el formato requerido la parseamos.
                    string CodigoBarras = CodigoLeido.Trim();
                    if (CodigoBarras.Length > 0)
                    {
                        string vin = CodigoLeido.Trim();
                        //Consultamos los datos para poder firmarnos en la base de datos.
                        #region Consulta de los datos para el Logueo en el Servidor Remoto
                        //ConexionBDchkDos objDB = new ConexionBDchkDos(this.CadenaConexion);
                        SqlConnection conBP = new SqlConnection();
                        SqlCommand bp_comand = new SqlCommand();

                        //conociendo el id_agencia procedemos a consultar los datos de conexion en la tabla transferencia
                        Q = "Select ip,usr_bd,pass_bd,nombre_bd,bd_alterna, dir_remoto_xml, dir_remoto_pdf,usr_remoto,pass_remoto, ip_almacen_archivos, smtpserverhost, smtpport, usrcredential, usrpassword ";
                        Q += " From SICOP_TRASMISION where id_agencia='" + NumeroSucursal + "'";

                        DataSet ds = this.objDB.Consulta(Q);
                        if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                        {
                            DataRow regConexion = ds.Tables[0].Rows[0];
                            string strconexionABussinesPro = string.Format("Data Source={0};Initial Catalog={1}; Persist Security Info=True; User ID={2};Password={3}", regConexion["ip"].ToString(), regConexion["bd_alterna"].ToString(), regConexion["usr_bd"].ToString(), regConexion["pass_bd"].ToString());

                            if (conBP.State.ToString().ToUpper().Trim() == "CLOSED")
                            {
                                DataSet ds1 = new DataSet();
                                try
                                {
                                    conBP.ConnectionString = strconexionABussinesPro;
                                    conBP.Open();
                                    // bp_comand.Connection = conBP;

                                    //teniendo la conexion con la base de datos vamos a consultar el campo del IdProspecto de Sicop
                                    Q = "SELECT PER_SICOP FROM ADE_VTAFI VT";
                                    Q += " INNER JOIN PER_PERSONAS PER ON PER.PER_IDPERSONA = VT.VTE_IDCLIENTE";
                                    Q += " WHERE VTE_TIPODOCTO = 'A' AND VTE_STATUS = 'I' AND VTE_SERIE = '" + CodigoBarras.Trim() + "'";

                                    System.Data.SqlClient.SqlDataAdapter objAdaptador = new System.Data.SqlClient.SqlDataAdapter(Q, conBP);
                                    objAdaptador.Fill(ds1, "Resultados");
                                    if (ds1.Tables.Count > 0)
                                    {
                                        if (ds1.Tables[0].Rows.Count > 0)
                                        {//no importa cuantos registros traiga siempre regresará solo la primer columna y del primer registro. 
                                            res = ds1.Tables[0].Rows[0][0].ToString().Trim();
                                        }
                                    }

                                }
                                catch (Exception ex1)
                                {
                                    res = "Error: Imposible conexion con BD de BP:" + ex1.Message;
                                }
                            }
                        }
                        else
                        {
                            res = "Error: No fue posible autenticarse en el servidor remoto";
                        }
                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    Utilerias.WriteToLog(ex.Message, "BuscaIDProspecto", Application.StartupPath + "\\Log.txt");
                }
                #endregion
            } //de que no es centralizada

            return res;        
        }

        
        #region procedimientos de soporte

        /*
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {                         
            switch (keyData)
            {
                case Keys.Enter:
                    string res = RegistrarSalida();
                        if (res.IndexOf("Error:") == -1)
                        {//no hubo error
                            DialogMensaje dialogo = new DialogMensaje("Atencion",res,false,"...",10,true,Color.Navy,Color.YellowGreen,"Paloma","");
                            dialogo.ShowDialog(); 
                        }
                        else
                        { //hubo un error.
                            DialogMensaje dial = new DialogMensaje("Atencion", res, false, "...", 10, true, Color.WhiteSmoke, Color.Red, "Error", "");
                            dial.ShowDialog();
                            this.txtTextoCodigoBarras.Text = "";
                        }
                    return true;
                default:
                    this.txtTextoCodigoBarras.Text += keyData.ToString();
                    return true;
            }
            
            return base.ProcessCmdKey(ref msg, keyData);

        }//ProcessCmdKey
        */


        public string MataProceso(string NombreProceso)
        {
            string res = "";
            try
            {
                if (NombreProceso.Trim() != "")
                {
                    NombreProceso = NombreProceso.Replace(".exe", "");
                    NombreProceso = NombreProceso.Replace(".EXE", "");

                    Process[] localByName = Process.GetProcessesByName(NombreProceso);
                    foreach (Process proceso in localByName)
                    {
                        proceso.CloseMainWindow();
                        if (proceso.HasExited == false)
                        {
                            proceso.Kill();
                            proceso.Close();
                            res = "El proceso: " + NombreProceso + " ha sido eliminado del TaskManager";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            return res;
        }

        /// <summary>
        /// Dada la ruta donde se encuentra un archivo ejecutable lanza su ejecucion
        /// </summary>
        /// <param name="rutaejecutable">El archivo ejecutable a ejecutar</param>
        /// <returns>Verdadero si pudo lanzar la ejecucion</returns>
        private bool LanzaEjecucion(string rutaejecutable)
        {
            bool res = false;
            try
            {

                //string filepath = @"C:\RepContavsNomina\Impersonate.bat";
                // Create the ProcessInfo object
                ProcessStartInfo psi = new ProcessStartInfo("cmd.exe");
                psi.UseShellExecute = false;
                psi.RedirectStandardOutput = true;
                psi.RedirectStandardInput = true;
                psi.RedirectStandardError = true;
                //impersonating
                //psi.UserName = "Administrator";
                //System.Security.SecureString psw = new SecureString();
                //foreach (Char ch in "Al3m4n14")
                //{
                //    psw.AppendChar(ch);
                //}
                //psi.Password = psw;
                //psi.Domain = System.Environment.MachineName;
                //psi.UseShellExecute = false;

                // Start the process           
                Process proc = Process.Start(psi);
                //StreamReader sr = File.OpenText(filepath);
                StreamWriter sw = proc.StandardInput;

                //while (sr.Peek() != -1)
                //{
                //    // Make sure to add Environment.NewLine for carriage return!
                //    sw.WriteLine(sr.ReadLine() + Environment.NewLine);
                //}
                sw.WriteLine(rutaejecutable + Environment.NewLine);

                //sr.Close();
                proc.Close();
                sw.Close();
                res = true;
            }
            catch (Exception ex)
            {
                Utilerias.WriteToLog(ex.Message, "LanzaEjecucion", Application.StartupPath + "\\Log.txt");
                Debug.WriteLine(ex.Message);
            }
            return res;
        }


        /// <summary>
        /// Consulta los procesos que se estan ejecutando en este momento 
        /// </summary>
        /// <param name="NombreProceso">A buscar si es que se está ejecutando</param>
        /// <returns>verdadero si el proceso está en ejecucion</returns>
        private bool EstaEnEjecucion(string NombreProceso)
        {
            bool res = false;
            try
            {
                if (NombreProceso.Trim() != "")
                {
                    NombreProceso = NombreProceso.Replace(".exe", "");
                    NombreProceso = NombreProceso.Replace(".EXE", "");

                    Process[] localByName = Process.GetProcessesByName(NombreProceso);
                    if (localByName.Length > 0)
                        res = true;
                    else
                        res = false;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            return res;
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

        #endregion

        

        

        public string ObtenDeArchivo(string RutaArchivo,string Que)
        {
            string res = "";                          
            string instruccion = "";
            FileStream fs = null;
            StreamReader sr = null;
            string cuote = "\"";

            Que = cuote + Que.Trim() + cuote;

            try
            {
                fs = new FileStream(RutaArchivo, FileMode.Open, FileAccess.ReadWrite);
                sr = new StreamReader(fs);
                
                int cont = 0;
                //int posicion = Que=="\"No_Vin\""?15:18;
                int posicion = 0;
                if (Que == "\"No_Vin\"")
                    posicion = 15;
                if (Que == "\"IdProspecto\"")
                    posicion = 18;
                if (Que == "\"TipoVenta\"")
                    posicion = 10;
                if (Que == "\"No_Factura\"")
                    posicion = 11;

                while (!sr.EndOfStream)
                {
                    instruccion = sr.ReadLine();
                    string[] Arr = instruccion.Split(',');
                    if (cont == 0)
                    {//el primer registro trae el nombre de las columnas buscamos en cual está el No_Vin.
                        if (Arr[posicion].Trim() != Que.Trim())
                        {
                            posicion = 0;
                            bool encontrado = false;
                            while (posicion < Arr.Length && !encontrado)
                            {
                                if (Arr[posicion].Trim() == Que.Trim())
                                {
                                    posicion--;
                                    encontrado = true;
                                }
                                posicion++;
                            }
                        }
                    }
                    else
                    {
                        res = Arr[posicion].Trim();
                        res = res.Replace("\"", ""); 
                    }
                    cont++;
                }
            }//del try
            catch (Exception ex)
            {
                Utilerias.WriteToLog("Error al buscar el vin en el archivo creado: " + ex.Message, "ObtenVinDeArchivo", Application.StartupPath + "\\Log.txt");
            }
            finally
            {
               if (sr != null)
                sr.Close();
               if (fs != null) 
                 fs.Close();
            }

            return res;
        }


        public string  procesaArchivoGeneradoporBPro(string ArchivoGenerado, string SoloNombre, string id_bitacora, string id_agencia)
        {
            string res = "";
            string Q = "";

            if (File.Exists(ArchivoGenerado))
            {         
                //El código escaneado debe ser leido del que contine el archivo txt, para que no se confunda cuando se escanea más de un código.
                string vinenarchivo = ObtenDeArchivo(ArchivoGenerado, "No_Vin");
                string facturaenarchivo = ObtenDeArchivo(ArchivoGenerado, "No_Factura");

                if (vinenarchivo.Trim() != "")
                {

                    string ArchivoRenombrado = SoloNombre.Trim();
                    ArchivoRenombrado = ArchivoRenombrado.ToUpper();
                    ArchivoRenombrado = ArchivoRenombrado.Replace(".TXT", "") + "_" + vinenarchivo.Trim() + "_" + facturaenarchivo.Trim() + ".TXT";

                    FileInfo fi = new FileInfo(ArchivoGenerado);

                    ArchivoRenombrado = fi.DirectoryName.Trim() + "\\" + ArchivoRenombrado.Trim();
                    if (File.Exists(ArchivoRenombrado))
                    {
                        Utilerias.WriteToLog("El archivo : " + ArchivoRenombrado + " ya existia se borra antes de agregarle el vin al generado por BPro", "fsw_Created", Application.StartupPath + "\\Log.txt");
                        File.Delete(ArchivoRenombrado); 
                    }

                    fi.MoveTo(ArchivoRenombrado); // le cambiamos el nombre agregandole al nombre del archivo el _vin y la factura.

                    string idprospenarchivo = ObtenDeArchivo(ArchivoRenombrado, "IdProspecto");
                    string tipoventa = ObtenDeArchivo(ArchivoRenombrado, "TipoVenta");
                    string factura = ObtenDeArchivo(ArchivoRenombrado, "No_Factura");                     
                    string tipo_auto = ""; //20200514 Seminuevo o Nuevo;

                    try
                    {
                        Q = "Select Top 1 * from SICOPCONFIGXMAQUINA where numero_sucursal = " + id_agencia.Trim();
                        DataSet ds = this.objDB.Consulta(Q);
                        foreach (DataRow registro in ds.Tables[0].Rows)
                        {
                            string strUsrRemoto = registro["usr"].ToString().Trim();  //this.Usr.Trim();
                            string strDominio = "";
                            string strIPFileStorage = registro["ip_remoto"].ToString().Trim();
                            string Pass = registro["pass"].ToString().Trim();
                            string CarpetaRemota = registro["carpeta_remota"].ToString().Trim();
                            string NumeroSucursal = registro["numero_sucursal"].ToString().Trim();
                            string strIPMaquina = registro["ip_local"].ToString().Trim();
                            string strNombreMaquina = registro["nombre"].ToString().Trim();
                            string strEnviar = registro["enviarcorreos"].ToString().Trim();

                            if (strUsrRemoto.IndexOf("\\") > -1)
                            {   // DANDRADE\sistemas     DANDRADE = dominio sistemas=usuario
                                strDominio = strUsrRemoto.Substring(0, strUsrRemoto.IndexOf("\\"));
                                strUsrRemoto = strUsrRemoto.Substring(strUsrRemoto.IndexOf("\\") + 1);
                            }

                            try
                            {
                                FileInfo Archivo = new FileInfo(ArchivoRenombrado);
                                ArchivoRenombrado = Archivo.Name.Trim();
                                string nuevaruta = CarpetaRemota + "\\" + ArchivoRenombrado.Trim();
                                string rutareal = ""; //ConsultaCarpetaDestino(idprospenarchivo.Trim()); //20200514
                                bool isIntercambio = false;
                                bool isFlotilla = false;
                                if (tipoventa.IndexOf("INTERCAMBIOS") >- 1) 
                                {
                                    isIntercambio = true;
                                    rutareal = Application.StartupPath + "\\Procesados\\INTERCAMBIOS"; //+ tipoventa.Trim();
                                }
                                if (tipoventa.IndexOf("FLOTILLA") > -1)
                                {
                                    isFlotilla = true;
                                    rutareal = Application.StartupPath + "\\Procesados\\FLOTILLAS"; //+ tipoventa.Trim();
                                }

                                //20200514
                                tipo_auto = facturaenarchivo.Substring(0, 1) == "A" ? "Nuevo" : "Seminuevo";

                                if (!isIntercambio && !isFlotilla) {
                                    rutareal = tipo_auto.Trim() == "Seminuevo" ? CarpetaRemota + "\\SEMINUEVOS" : "";
                                }      

                                nuevaruta = rutareal.Trim() == "" ? nuevaruta.Trim() : rutareal.Trim() + "\\" + ArchivoRenombrado.Trim();

                                if (File.Exists(nuevaruta))
                                    File.Delete(nuevaruta);

                                Archivo.CopyTo(nuevaruta); //Esta linea es la que envia el archivo a la carpeta de SICOP.

                                if (File.Exists(nuevaruta))
                                {
                                    Utilerias.WriteToLog("Se Envió el archivo a la carpeta de SICOP : " + nuevaruta, "fsw_Created", Application.StartupPath + "\\Log.txt");
                                    res = nuevaruta;
                                }

                                if (File.Exists(Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim()))
                                    File.Delete(Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim());

                                //Archivo.CopyTo(Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim());
                                //20230102 Siempre copia a la carpeta de \\Procesados sin embargo hay ocasiones en que lo QUITA de la ruta de SICOP.
                                //se comentaria la linea de arriba y se cambia con un nuevo Archivo para copiar a \\Procesados 
                                FileInfo ArchivoProcesados = new FileInfo(nuevaruta); //lo tomamos de la carpeta de SICOP.
                                ArchivoProcesados.CopyTo(Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim());
                                if (File.Exists(Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim()) && File.Exists(nuevaruta))
                                { Utilerias.WriteToLog("Se copio a Procesados  y se conserva en:  " + nuevaruta, "ProcesaArchivoGeneradoporBPro", Application.StartupPath + "\\Log.txt"); }
                                else
                                { Utilerias.WriteToLog("Faltá en una de las 2 carpetas: \\Procesados o en " + nuevaruta, "ProcesaArchivoGeneradoporBPro", Application.StartupPath + "\\Log.txt"); }

                                //20230102 ya no se borra el archivo renombrado + vin + factura, el cual se va a ir quedando en cada una de las carpetas E:\LB\PROGRAMAS\RegistraSalida\??????Recibidos\VentasBPro\ 
                                //Esto para garantizar que no es esta línea la que borra el archivo en la carpeta SEEKOP
                                //if (File.Exists(Archivo.FullName))
                                //{
                                    //Utilerias.WriteToLog("Se borra el archivo: " + Archivo.FullName, "fsw_Created", Application.StartupPath + "\\Log.txt");
                                //    Archivo.Delete();
                                //}                                
                            }
                            catch (Exception exe)
                            {
                                Debug.WriteLine(exe.Message);
                                Utilerias.WriteToLog("Error: \r" + exe.Message + "\r", "fsw_Created", Application.StartupPath + "\\Log.txt");
                            }                            
                        }//de que encontró los parametros de envio.
                    }
                    catch (Exception ex)
                    {
                        Utilerias.WriteToLog("Error al loguearse  en la carpeta remota \n\r" + ex.Message, "fsw_Created", Application.StartupPath + "\\Log.txt");
                        Debug.WriteLine(ex.Message);
                    }
                    //termina envio                     
                }//de que se pudo recuperar el vin del archivo.
                else
                {
                    Utilerias.WriteToLog("No se pudo recuperar el vin del archivo: " + ArchivoGenerado, "fsw_Created", Application.StartupPath + "\\Log.txt");
                }
            }
            else{
              Utilerias.WriteToLog("* El archivo: " + ArchivoGenerado + " no existe", "procesaArchivoGeneradoporBPro", Application.StartupPath + "\\Log.txt");
            }
            return res;
        }

        private string ConsultaCarpetaDestino(string idsicop)
        { string res="";
            if (idsicop.Trim() != "")
            {
                idsicop = idsicop.Substring(0, 6);
                string Q = "Select carpeta_local from SICOP_PREFIJOSIDSICOP where prefijo='" + idsicop + "'";
                res = this.objDB.ConsultaUnSoloCampo(Q).Trim(); 
            }
            return res;        
        }


        private void fsw_Created(object sender, FileSystemEventArgs e)
        {
            Utilerias.WriteToLog(" Se creó el archivo : " + e.FullPath, "fsw_Created", Application.StartupPath + "\\Log.txt");
            //procesaArchivoGeneradoporBPro(e.FullPath,e.Name);           
        }

        public string EnviaCorreo(string vin, string idsicop, string rutaarchivocreado, string desde, string Id_Maquina, string NumeroSucursal, string strIPMaquina,string NombreMaquina,string Enviar,string TipoVenta, string CodigoErrorSICOP, string id_bitacora)
        {
            string res = "";
            try
            {
                if (Enviar.Trim() == "True" || Enviar.ToUpper()=="TRUE")
                {
                string NombreAgencia = this.objDB.ConsultaUnSoloCampo("Select alias from AGENCIAS where id_agencia='" + NumeroSucursal.Trim() + "'"); 
                string Q = "Select * from SICOP_TRASMISION where id_agencia = '" + NumeroSucursal.Trim() + "'"; 
                DataSet ds = this.objDB.Consulta(Q);
                foreach (DataRow registro in ds.Tables[0].Rows)
                {                    
                    string smtpserverhost = registro["smtpserverhost"].ToString().Trim();
                    string smtpport = registro["smtpport"].ToString().Trim();
                    string usrcredential = registro["usrcredential"].ToString().Trim();
                    string usrpassword = registro["usrpassword"].ToString().Trim();
                    string EnableSsl = registro["enable_ssl"].ToString().Trim();
                    string plantillaHTML = registro["plantillaHTML"].ToString().Trim();

                        
                        Utilerias.WriteToLog("Intento de envio de correo desde: " + desde.Trim() + " vin: " + vin + " idsicop: " + idsicop.Trim() + " archivo: " + rutaarchivocreado + " id_maquina: " + Id_Maquina, "EnviaCorreo", Application.StartupPath + "\\Log.txt");

                        string rutaplantilla = Application.StartupPath;
                        rutaplantilla += "\\" + plantillaHTML.Trim();
                        clsEmail correoLog = new clsEmail(smtpserverhost.Trim(), Convert.ToInt16(smtpport), usrcredential.Trim(), usrpassword.Trim(), EnableSsl.Trim());
                        MailMessage mensaje = new MailMessage();
                        mensaje.Priority = System.Net.Mail.MailPriority.Normal;
                        mensaje.IsBodyHtml = false;

                        string str_subject = ""; 
                    if (idsicop.Trim() == "")
                             str_subject = "Interfaz: " + NombreAgencia.Trim() + " V.I.N. : " + vin.Trim() + " Sin id prospecto de SICOP ";
                        else
                             str_subject = "Interfaz: " + NombreAgencia.Trim() + " V.I.N. : " + vin.Trim() + " Id Prospecto: " + idsicop.Trim();
                    
                    if (CodigoErrorSICOP.Trim() != "")
                    {
                        str_subject += " Error en carga de SICOP";
                    }
                   
                    mensaje.Subject = str_subject.Trim();

                        string Remitente = "Sistemas de Grupo Andrade";
                        
                        Q = "Select Isnull(fecha_promesa_pago,'01/01/1900') From SICOP_BITACORA where id_bitacora=" + id_bitacora.Trim();
                        string Fecha_Promesa_Pago = this.objDB.ConsultaUnSoloCampo(Q);
                        Q = "Select isnull(tipo_auto,'') From SICOP_BITACORA where id_bitacora=" + id_bitacora.Trim(); 
                        string tipo_auto = this.objDB.ConsultaUnSoloCampo(Q);
                        string campoconsultar = "";

                        if (tipo_auto.Trim() == "Seminuevo") //20200628
                        {
                            //Si trae idprospecto se envia a Desarrollo, si no trae se envia a Operacion
                            campoconsultar = idsicop == "" ? "i_operacion_semi" : "i_desarrollo_semi";
                            Q = "Select " + campoconsultar + " from SICOPCONFIGXMAQUINA where id_maquina=" + Id_Maquina;
                        }
                        else
                        {
                            //Si trae idprospecto se envia a Desarrollo, si no trae se envia a Operacion
                            campoconsultar = idsicop == "" ? "interesados_operacion" : "interesados_desarrollo";
                            Q = "Select " + campoconsultar + " from SICOPCONFIGXMAQUINA where id_maquina=" + Id_Maquina;
                        }

                        string emailsavisar = this.objDB.ConsultaUnSoloCampo(Q);
                        if (emailsavisar.Trim() != "")
                        {
                            string[] EmailsEspeciales = emailsavisar.Split(',');

                            foreach (string Email in EmailsEspeciales)
                            {
                                mensaje.To.Add(new MailAddress(Email.Trim()));
                            }

                            mensaje.From = new MailAddress(usrcredential.Trim(), Remitente.Trim());

                            if (rutaarchivocreado.Trim() != "")
                                mensaje.Attachments.Add(new Attachment(rutaarchivocreado));

                            string cadenaLog = "Se ha escaneado el codigo de barras del siguiente Número de Serie : " + vin.Trim() + "\n" + "\r";
                                   cadenaLog += "Desde la agencia: " + NombreAgencia.Trim() + "\n" + "\r";

                            if (tipo_auto.Trim() != "")
                            {
                              cadenaLog += "\n" + "\r";
                              cadenaLog += " Tipo Auto : " + tipo_auto.ToUpper().Trim();
                              cadenaLog += "\n" + "\r";
                            }

                            if (Fecha_Promesa_Pago.Trim() != "" && Fecha_Promesa_Pago != "01/01/1900 12:00:00 a. m." && Fecha_Promesa_Pago != "01/01/1900")
                             {
                                       cadenaLog += "\n" + "\r";
                                       cadenaLog += " Fecha Promesa de Pago: " + Fecha_Promesa_Pago.Trim();
                                       cadenaLog += "\n" + "\r";                            
                             }

                            if (CodigoErrorSICOP.Trim() != "")
                            {
                                cadenaLog += " ERROR AL REGISTRAR EL ARCHIVO DE VENTAS EN SICOP: " + "\n" + "\r";
                                cadenaLog += " " + CodigoErrorSICOP.Trim() + "\n" + "\r";
                                cadenaLog += " POR TAL MOTIVO LA INFORMACION RESIDENTE EN EL ARCHIVO DE VENTAS NO SE VERA REFLEJADA EN SICOP" + "\n" + "\r";
                                cadenaLog += " [se adjunta el archivo generado]" + "\n" + "\r";
                            }
                            
                            cadenaLog += " TIPO VENTA: " + TipoVenta.Trim() + "\n" + "\r";
                            if (TipoVenta.Trim() == "INTERCAMBIOS")
                            {
                                cadenaLog += " LOS INTERCAMBIOS NO SE REGISTRAN COMO VENTAS EN SICOP " + "\n" + "\r";                                
                                cadenaLog += " [se adjunta el archivo generado]" + "\n" + "\r";
                            }
                     
                            cadenaLog += "\n" + "\r";
                            if (idsicop.Trim() == "")
                            {                                
                                cadenaLog += " NO SE TIENE REGISTRADO UN ID PROSPECTO DE SICOP " + "\n" + "\r";
                                cadenaLog += " POR TAL MOTIVO LA INFORMACION RESIDENTE EN EL ARCHIVO DE VENTAS NO SE VERA REFLEJADA EN SICOP" + "\n" + "\r";
                                cadenaLog += " [se adjunta el archivo generado]" + "\n" + "\r";
                            }
                            else
                            {
                                cadenaLog += " Id prospecto de SICOP: " + idsicop.Trim() + "\n" + "\r";
                            }
                            cadenaLog += "\n" + "\r";
                            cadenaLog += " Datos del cliente de escaneo: " + "\n" + "\r";
                            cadenaLog += " Maquina: " + NombreMaquina.Trim() + "\n" + "\r";
                            cadenaLog += " Ip Local: " + strIPMaquina.Trim() + "\n" + "\r";
                            cadenaLog += " # Sucursal: " + NumeroSucursal + "\n" + "\r";

                            //mensaje.Body = cadenaLog.Trim();

                            Dictionary<string, string> TextoIncluir = new Dictionary<string, string>();

                            TextoIncluir.Add("fecha", DateTime.Now.ToString("dd-MM-yyyy"));
                            TextoIncluir.Add("hora", DateTime.Now.ToString("HH:mm:ss"));
                            TextoIncluir.Add("LogEjecucion", cadenaLog);

                            //AlternateView vistaplana = AlternateView.CreateAlternateViewFromString(correo.CreaCuerpoPlano(TextoIncluir), null, "text/plain");
                            AlternateView vistahtml = AlternateView.CreateAlternateViewFromString(correoLog.CreaCuerpoHTML(rutaplantilla, TextoIncluir).ToString(), null, "text/plain");

                            //LinkedResource logo = new LinkedResource(rutalogo);
                            //logo.ContentId = "companylogo";
                            //vistahtml.LinkedResources.Add(logo);

                            //mensaje.AlternateViews.Add(vistaplana);
                            mensaje.AlternateViews.Add(vistahtml);
                            correoLog.MandarCorreo(mensaje);
                            res = "Envio exitoso del Log";
                        } //de si hay cuentas de correo a quien enviar.                    
                    }//de que pudo traer los datos de envio.
                } //De que si se envian los correos
            }
            catch (Exception ex)
            {
                res = ex.Message;
                Utilerias.WriteToLog(ex.Message, "EnviaCorreo", Application.StartupPath + "\\Log.txt");
            }
            return res;
        }

        #region Persistencia
        

        //le debe llegar sin la extension .exe
        private int CuentaInstancias(string NombreProceso)
        {
            int res = 0;
            try
            {
                if (NombreProceso.Trim() != "")
                {
                    Process[] localByName = Process.GetProcessesByName(NombreProceso);
                    res = localByName.Length;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            return res;
        }

        #endregion

        private void timerThread_Tick(object sender, EventArgs e)
        {            
            this.ntiBalloon.ShowBalloonTip(1, "SICOP REGISTRO FACTURAS", " Leyendo Tabla Intermedia", ToolTipIcon.Info);
               ProcesaBitacora();
        }

        public void CorrigeRegistrosSinIdAgencia()
        {

            string Q = " Select sb.id_bitacora,fecha, sb.quien, cxm.id_maquina, cxm.numero_sucursal from SICOPCONFIGXMAQUINA cxm, SICOP_BITACORA sb";
                   Q += " where cxm.id_maquina= substring(sb.quien,CHARINDEX('=',sb.quien)+1, CHARINDEX(' ',sb.quien)-CHARINDEX('=',sb.quien)-1)";
                   Q += " and sb.id_bitacora>28879";
                   Q += " and sb.id_agencia=''";
                   try
                   {
                       DataSet ds = objDB.Consulta(Q);
                       if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                       {
                           foreach (DataRow reg in ds.Tables[0].Rows)
                           {
                               Utilerias.WriteToLog("Se le pone id_agencia=" + reg["numero_sucursal"].ToString().Trim() + " al id_bitacora=" + reg["id_bitacora"].ToString().Trim(), "CorrigeRegistrosSinIdAgencia", Application.StartupPath + "\\Log.txt");
                           }

                           Q = "update  SICOP_BITACORA  set id_agencia=cxm.numero_sucursal ";
                           Q += " From SICOPCONFIGXMAQUINA cxm";
                           Q += " where cxm.id_maquina = substring(quien,CHARINDEX('=',quien)+1, CHARINDEX(' ',quien)-CHARINDEX('=',quien)-1)";
                           Q += " and id_agencia=''";
                           Q += " and id_bitacora >28879";
                           this.objDB.EjecUnaInstruccion(Q);
                       }
                   }
                   catch (Exception ex)
                   {
                       Utilerias.WriteToLog(ex.Message + " Q=" + Q, "CorrigeRegistrosSinIdAgencia", Application.StartupPath + "\\Log.txt");
                   }
        }


        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            string NombreEjecutable = Application.ExecutablePath.Trim();
            NombreEjecutable = NombreEjecutable.Replace(Application.StartupPath.Trim(),""); 
            NombreEjecutable = NombreEjecutable.Replace("\\","");

            this.ntiBalloon.Icon = this.Icon;
            this.ntiBalloon.Text = "SICOP REGISTRO DE FACTURAS: " + NombreEjecutable;
            this.ntiBalloon.Visible = true;
            this.ntiBalloon.ShowBalloonTip(1, "SICOP REGISTRO DE FACTURAS " + NombreEjecutable, " En espera de instrucciones ", ToolTipIcon.Info);
            this.toolStripMenuItem1.Text = "CERRAR " + NombreEjecutable.Trim();            
            this.Hide();
            this.Visible = false;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.ntiBalloon.Visible = false;
            this.ntiBalloon = null;
        }

        /// <summary>
        /// Estando dentro de un hilo que se ejecuta cada minuto en un ciclo de hasta 10 ocasiones, se consulta en la base de datos correspondiente en SICOP el estatus del procesamiento de ese archivo, se consulta el log en SICOP por el número de factura
        /// El hilo se debe autodestruir en cuanto se envie el correo, ya sea con error o bien con Exito en el proceso.
        /// </summary>
        /// <param name="facturaenarchivo"></param>
        /// <param name="vinenarchivo"></param>
        /// <param name="idprospenarchivo"></param>
        /// <param name="?"></param>
        /// <param name="proceso"></param>
        /// <param name="id_maquina"></param>
        /// <param name="NumeroSucursal"></param>
        /// <param name="strIPMaquina"></param>
        /// <param name="strNombreMaquina"></param>
        /// <param name="strEnviar"></param>
        /// <param name="tipoventa"></param>
        private void SensaResultadoCargaEnSicop(string id_bitacora, string facturaenarchivo,string vinenarchivo,string idprospenarchivo,string RutaArchivoGenerado, string proceso,string id_maquina,string NumeroSucursal,string strIPMaquina,string strNombreMaquina,string strEnviar,string tipoventa)
        {
            bool Consulta = true;
            int veces=0;
            string Q="";
            int TopeIntentos = Convert.ToInt16(this.TopeMinBusqenBDSicop);

            try
            {

                Q = "Select conexion_sicop from SICOPCONFIGXMAQUINA WHERE id_maquina=" + id_maquina;
                string conexion = this.objDB.ConsultaUnSoloCampo(Q);

                ConexionBD objDBSicop = new ConexionBD(conexion.Trim());

                while (Consulta == true && veces < TopeIntentos)
                {

                    Q = "Update SICOP_BITACORA set ";
                    Q += " fh_busqueda_sicop=getdate()";
                    Q += " where id_bitacora={0}";
                    Q = string.Format(Q, id_bitacora.Trim());
                    this.objDB.EjecUnaInstruccion(Q);

                    Q = " Select top 1 Cancelacion,Codigo,rtrim(ltrim(Mensaje)) as Mensaje";
                    Q += " From LogEnvioInterfazWS";
                    Q += " where idVenta='" + facturaenarchivo.Trim() + "'";
                    DataSet ds = objDBSicop.Consulta(Q);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow reg in ds.Tables[0].Rows)
                        {
                            string codigo_sicop = reg["Codigo"].ToString().Trim();
                            string mensaje_sicop = reg["Mensaje"].ToString().Trim();
                            string resultado_busqueda = codigo_sicop == "20" ? "Encontrado con error" : "Encontrado ok: " + facturaenarchivo;

                            //Codigo = 20 es un error de registro en SICOP enviamos el correo de Confirmacion con Código de Error.                      
                            if (codigo_sicop == "20")
                            {
                                EnviaCorreo(vinenarchivo.Trim(), idprospenarchivo, RutaArchivoGenerado, proceso, id_maquina, NumeroSucursal, strIPMaquina, strNombreMaquina, strEnviar, tipoventa.Trim(), "20: " + mensaje_sicop.Trim(), id_bitacora.ToString());
                            }
                            else
                            {
                                EnviaCorreo(vinenarchivo.Trim(), idprospenarchivo, RutaArchivoGenerado, proceso, id_maquina, NumeroSucursal, strIPMaquina, strNombreMaquina, strEnviar, tipoventa.Trim(), "",id_bitacora.ToString());
                            }

                            Q = "Update SICOP_BITACORA set cancelacion_sicop='{0}',";
                            Q += " codigo_sicop='{1}',";
                            Q += " mensaje_sicop='{2}',";
                            Q += " fh_busqueda_sicop=getdate(),";
                            Q += " resultado_busqueda='{3}'";
                            Q += " where id_bitacora={4}";
                            Q = string.Format(Q, reg["Cancelacion"].ToString().Trim(), codigo_sicop.Trim(), mensaje_sicop.Trim(), resultado_busqueda.Trim(), id_bitacora.Trim());
                            this.objDB.EjecUnaInstruccion(Q);

                            Consulta = false;
                            veces = 500;
                            //De una vez detenemos el hilo.
                            #region mata el hilo
                            try
                            {
                                if (this.dicHilos.ContainsKey(vinenarchivo.Trim()))
                                {

                                    if (this.dicHilos[vinenarchivo.Trim()] != null)
                                    {
                                        this.dicHilos[vinenarchivo.Trim()].Interrupt();
                                        this.dicHilos[vinenarchivo.Trim()] = null;
                                    }
                                    // if (!this.dicHilos[Folio_Operacion].IsAlive)
                                    this.dicHilos.Remove(vinenarchivo.Trim());
                                }
                            }
                            catch (Exception ex3)
                            {
                                Debug.WriteLine(ex3.Message);
                                Utilerias.WriteToLog(ex3.Message, "SensaResultadoCargaEnSicop", Application.StartupPath + "\\Log.txt");
                            }
                            #endregion
                        }
                    }
                    else
                    {
                     //Utilerias.WriteToLog("No encontró registro: " + Q.Trim() + " " + objDBSicop.DameCadenaConexion, " SensaResultadoCargaEnSicop", Application.StartupPath + "\\Log.txt");
                     Q = "Update SICOP_BITACORA set ";
                     Q += " resultado_busqueda='{0}'";
                     Q += " where id_bitacora={1}";
                     Q = string.Format(Q, "Aun no encontrado intento: " + veces.ToString() + " de " + TopeIntentos.ToString(),id_bitacora.Trim());
                     this.objDB.EjecUnaInstruccion(Q);
                    }
                    Thread.Sleep(60000);
                    veces++;
                }//Del while de la consulta.   

                //Por si nunca encontró en la BD esa factura siempre mata el hilo.
                if (veces >= TopeIntentos || Consulta == false)
                {
                    #region mata el hilo
                    try
                    {
                        if (this.dicHilos.ContainsKey(vinenarchivo.Trim()))
                        {

                            if (this.dicHilos[vinenarchivo.Trim()] != null)
                            {
                                this.dicHilos[vinenarchivo.Trim()].Interrupt();
                                this.dicHilos[vinenarchivo.Trim()] = null;
                            }
                            // if (!this.dicHilos[Folio_Operacion].IsAlive)
                            this.dicHilos.Remove(vinenarchivo.Trim());
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                        Utilerias.WriteToLog(ex.Message, "SensaResultadoCargaEnSicop", Application.StartupPath + "\\Log.txt");
                    }
                    #endregion
                }


            }
            catch (Exception ex2)
            {
                Utilerias.WriteToLog( ex2.Message,"SensaResultadoCargaEnSicop ex2", Application.StartupPath + "\\Log.txt");
            }
        }

        private void timerReproceso_Tick(object sender, EventArgs e)
        {
            int hora = DateTime.Now.Hour;

            if (hora == 16 || hora == 19 || hora == 21)
            {
                string Q = "select id_bitacora from SICOP_BITACORA";
                       Q += " where resultado_busqueda like 'Aun no encontrado intento:%de " + this.TopeMinBusqenBDSicop.Trim()  +"'";
                       Q += " and Convert(char(8),fh_creacion_av,112) = Convert(char(8),getdate(),112)";
                       Q += " and isnull(mensaje_sicop,'')=''";

                       DataSet ds = this.objDB.Consulta(Q);
                       if ( ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count>0)
                       {
                           Utilerias.WriteToLog("Hay elementos ha reprocesar : " + ds.Tables[0].Rows.Count.ToString(), "timerReproceso_Tick", Application.StartupPath + "\\Log.txt");
                           
                           foreach (DataRow reg in ds.Tables[0].Rows)
                           {                                                              
                               Q = " Update SICOP_BITACORA set intentos=0,factura=null,cancelacion_sicop=null,codigo_sicop=null,";
                               Q += " mensaje_sicop=null,fh_busqueda_sicop=null,resultado_busqueda=null,fh_creacion_av=null,fh_envio_bp=null";
                               Q += " where id_bitacora =  " + reg["id_bitacora"].ToString().Trim();
                               if (this.objDB.EjecUnaInstruccion(Q) > 0 )
                                   Utilerias.WriteToLog("Reprocesando id_bitacora: " + reg["id_bitacora"].ToString().Trim(), "timerReproceso_Tick", Application.StartupPath + "\\Log.txt");
                           } //del for
                           this.ntiBalloon.ShowBalloonTip(1, "SICOP CODIGOS DE BARRAS SALIDAS", " RE-Procesando Bitacora ", ToolTipIcon.Info);
                           //ProcesaBitacora(); //Lo hace cada minuto, para que estresarlo más?
                       }//de si hay algo por reprocesar
            }//de la hora de reproceso.
        }


        #region EnviarCorreoAvisoProblema
        public string EnviaCorreoSISTEMAS(string vin, string NumeroSucursal,string id_bitacora, string TipoAuto, string NombreBase, string Query)
        {
            string res = "";
            try
            {
                    string NombreAgencia = this.objDB.ConsultaUnSoloCampo("Select alias from AGENCIAS where id_agencia='" + NumeroSucursal.Trim() + "'");
                    string Q = "Select * from SICOP_TRASMISION where id_agencia = '" + NumeroSucursal.Trim() + "'";
                    DataSet ds = this.objDB.Consulta(Q);
                    foreach (DataRow registro in ds.Tables[0].Rows)
                    {
                        string smtpserverhost = registro["smtpserverhost"].ToString().Trim();
                        string smtpport = registro["smtpport"].ToString().Trim();
                        string usrcredential = registro["usrcredential"].ToString().Trim();
                        string usrpassword = registro["usrpassword"].ToString().Trim();
                        string EnableSsl = registro["enable_ssl"].ToString().Trim();
                        string plantillaHTML = registro["plantillaHTML"].ToString().Trim();


                        Utilerias.WriteToLog("Intento de envio de correo desde: EnviaCorreoSISTEMAS vin: " + vin , "EnviaCorreoSISTEMAS", Application.StartupPath + "\\Log.txt");

                        string rutaplantilla = Application.StartupPath;
                        rutaplantilla += "\\" + plantillaHTML.Trim();
                        clsEmail correoLog = new clsEmail(smtpserverhost.Trim(), Convert.ToInt16(smtpport), usrcredential.Trim(), usrpassword.Trim(), EnableSsl.Trim());
                        MailMessage mensaje = new MailMessage();
                        mensaje.Priority = System.Net.Mail.MailPriority.Normal;
                        mensaje.IsBodyHtml = false;

                        string str_subject = "No se ha actualizado la fecha de Entrega en UNI_PEDIUNI o USN_PEDIDO " + vin.Trim();

                        mensaje.Subject = str_subject.Trim();

                        string Remitente = "Sistemas de G. Andrade";
                        
                        string emailsavisar = "luis.bonnet@grupoandrade.com";
                        if (emailsavisar.Trim() != "")
                        {
                            string[] EmailsEspeciales = emailsavisar.Split(',');

                            foreach (string Email in EmailsEspeciales)
                            {
                                mensaje.To.Add(new MailAddress(Email.Trim()));
                            }

                            mensaje.From = new MailAddress(usrcredential.Trim(), Remitente.Trim());

                            //if (rutaarchivocreado.Trim() != "")
                              //  mensaje.Attachments.Add(new Attachment(rutaarchivocreado));

                            string cadenaLog = "Se ha escaneado el codigo de barras del siguiente Número de Serie : " + vin.Trim() + "\n" + "\r";
                            cadenaLog += "Desde la agencia: " + NombreAgencia.Trim() + "\n" + "\r";

                            if (TipoAuto.Trim() != "")
                            {
                                cadenaLog += "\n" + "\r";
                                cadenaLog += " Tipo Auto : " + TipoAuto.ToUpper().Trim();
                                cadenaLog += "\n" + "\r";
                            }
                                cadenaLog += "\n" + "\r";

                                cadenaLog += " BASE DE DATOS " + NombreBase.Trim() + "\n" + "\r";
                                cadenaLog += " Query: " + Query.Trim() + "\n" + "\r";                                


                            cadenaLog += "\n" + "\r";

                            //mensaje.Body = cadenaLog.Trim();

                            Dictionary<string, string> TextoIncluir = new Dictionary<string, string>();

                            TextoIncluir.Add("fecha", DateTime.Now.ToString("dd-MM-yyyy"));
                            TextoIncluir.Add("hora", DateTime.Now.ToString("HH:mm:ss"));
                            TextoIncluir.Add("LogEjecucion", cadenaLog);

                            //AlternateView vistaplana = AlternateView.CreateAlternateViewFromString(correo.CreaCuerpoPlano(TextoIncluir), null, "text/plain");
                            AlternateView vistahtml = AlternateView.CreateAlternateViewFromString(correoLog.CreaCuerpoHTML(rutaplantilla, TextoIncluir).ToString(), null, "text/plain");

                            //LinkedResource logo = new LinkedResource(rutalogo);
                            //logo.ContentId = "companylogo";
                            //vistahtml.LinkedResources.Add(logo);

                            //mensaje.AlternateViews.Add(vistaplana);
                            mensaje.AlternateViews.Add(vistahtml);
                            correoLog.MandarCorreo(mensaje);
                            res = "Envio exitoso del Log";
                        } //de si hay cuentas de correo a quien enviar.                    
                    }//de que pudo traer los datos de envio.                
            }
            catch (Exception ex)
            {
                res = ex.Message;
                Utilerias.WriteToLog(ex.Message, "EnviaCorreoSISTEMAS", Application.StartupPath + "\\Log.txt");
            }
            return res;
        }

        #endregion
                
        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            string NombreEjecutable = Application.ExecutablePath.Trim();
            NombreEjecutable = NombreEjecutable.Replace(Application.StartupPath.Trim(), "");
            NombreEjecutable = NombreEjecutable.Replace("\\", "");

            try
            {
                if (e.ClickedItem == toolStripMenuItem1)  //toolStripMenuItem1
                {
                    DialogResult res = MessageBox.Show(this,  "Confirme desea cerrar la aplicación: " + NombreEjecutable.Trim(),"Atención", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                    if (res == DialogResult.Yes)
                    {
                        Utilerias.WriteToLog("Usuario seleccionó cerrar el aplicativo", "contextMenuStrip1_ItemClicked", Application.StartupPath + "\\Log.txt");
                        Application.Exit();  
                    }
                }
            }
            catch (Exception ex)
            {
                Utilerias.WriteToLog(ex.Message, "contextMenuStrip1_ItemClicked", Application.StartupPath + "\\Log.txt");
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }


        public void CreaArchivoCancelacion(string LineaEncabezado, string LineaDetalle, string rutaArchivo)
        {
            FileStream fs = null;
            StreamWriter sw = null;
            try
            {
                //si el archivo ya existe lo va a sobreescribir
                    fs = new FileStream(rutaArchivo, FileMode.Create, FileAccess.Write, FileShare.None);
                    sw = new StreamWriter(fs, Encoding.UTF8);
                    if (LineaEncabezado.Trim() != "")
                        sw.WriteLine(LineaEncabezado.Trim());
                    if (LineaDetalle.Trim() != "")
                        sw.WriteLine(LineaDetalle.Trim());
                    sw.Close();
                    fs.Close();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            finally
            {
                if (sw != null)
                {
                    sw.Close();
                    sw = null;
                }
                if (fs != null)
                {
                    fs.Close();
                    fs = null;
                }
            }
        }



    }
}
