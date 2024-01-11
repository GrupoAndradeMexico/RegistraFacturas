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
using System.Net;
using System.Net.Security;

using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Net.NetworkInformation;

using System.Management;
using Microsoft.Win32;
using System.Net;
using System.Net.Mail;

namespace RegistraSalida
{
    public partial class Form1 : Form
    {
        /*     ************************** ESTA ES LA VERSION QUE CORRE EN EL SERVIDOR *************************************
        
        29/09/2014
        
        La interfaze de BPro únicamente emitirá el archivo de Ventas (el asociado al vin que se le pase), en el momento del scaneo de la pistola.
        Este archivo hay que agregarle el vin al final del nombre del archivo.
        Y pasarlo a la ruta de la carpeta de recepcion.
         
        Tanto los vendedores y los cancelaciones se harán en un proceso interno en BPro x separado.
        
        
        20150203 Se requiere que al momento de scanear verifique que exista un id_prospecto de sicop en la base de datos asociado al vin que se está
         * escaneando.
         * Si tiene id_prospecto de sicop, únicamente se dará aviso a los interesados de desarrollo via correo electrónico, adjuntando el archivo txt genereado y enviado al 43
         * Si no tiene id_prospecto de sicop, únicamente se dará aviso a los interesados de operacion via correo electrónico. 
        
         20150213 Correra en el servidor cada X minutos.
         * Revisará cada maquina en la tabla SICOPCONFIGXMAQUINA por los archivos yyyyMMddHHmmss_vin.txt cuando el cliente no tiene conexion.
         * Traerá al servidor dichos archivos y los registrará en la tabla SICOP_BITACORA para su posterior ejecucion.
         * Buscara en la tabla SICOP_BITACORA por aquellos registros que no hayan sido procesados y mandara a ejecutar BPro.
         * Finalmente transladará los archivos a la carpeta de SICOP.
         * 
         * Al sensar el fsw que se BPro ha creado un nuevo archivo, debe decidir a la carpeta de cual agencia debe enviar el archivo en base a la nomenclatura definida en la tabla: SICOP_PREFIJOSIDSICOP
         * pues el IDSICOP no es único en SICOP es decir, el id sicop 500 en FZA Area es para LBonnet mientras que el id sicop 500 en Zaragoza es para Fulanito. 
         
         * 20150401 Cuando el TipoVenta es INTERCAMBIOS, no se envía a la carpeta de SICOP, se envia a una carpeta llamada INTERCAMBIOS, se trata como si no trajera ID SICOP. 
         * 
         * 20150429 Para poder implementar todas las agencias, hay que parametrizar en base de datos: RutaEjecutableBPro, DirectorioArchivosSICOP y dinámicamente crear un fsw con el directorio donde deja los archivos de Ventas BPro.
         * a fin de que si se llega a trabar la interfaz, esto únicamente sea para la agencia con problemas.
         * 
         * 20150505 No se está utilizando el fsw_Created.
         */

        // "C:\AndradeGPO\ActualizarCampoEnBP\Ejecutable\BusinessProSICOP.exe" SICOP GMI GAZM_Zaragoza Exporta C:\AndradeGPO\ActualizarCampoEnBP\SiCoP\Generar\ parametro_ocioso.txt 3N1CK3CD9DL259265 1000 5063

        string ConnectionString = System.Configuration.ConfigurationSettings.AppSettings["ConnectionString"];
        ConexionBD objDB = null;
        
        //20150429 string RutaEjecutableBPro = System.Configuration.ConfigurationSettings.AppSettings["RutaEjecutableBPro"];
        //20150429 string DirectorioArchivosSICOP = System.Configuration.ConfigurationSettings.AppSettings["DirectorioArchivosSICOP"]; = carpeta_local_ventas
        //20150429 string Mascara = System.Configuration.ConfigurationSettings.AppSettings["Mascara"];


        string Latencia = System.Configuration.ConfigurationSettings.AppSettings["Latencia"];
        string MinutosEsperaraBPro = System.Configuration.ConfigurationSettings.AppSettings["MinutosEsperaraBPro"];
        string TotalIntentos = System.Configuration.ConfigurationSettings.AppSettings["TotalIntentos"];

        //ArrayList fw = new ArrayList();


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
            string Q = "Select * from SICOPCONFIGXMAQUINA where activo='True' ";
            //Q += " and id_maquina=3";
            Q += " order by numero_sucursal";

            DataSet ds = this.objDB.Consulta(Q);
            foreach (DataRow lector in ds.Tables[0].Rows)
            {
                try
                {
                    string id_agencia = lector["numero_sucursal"].ToString().Trim();
                    string UsuarioBPRo1 = lector["usuario_bpro"].ToString().Trim();
                    string BDBPRo1 = lector["bd_bpro"].ToString().Trim();
                    string carpeta_local_ventas = lector["carpeta_local_ventas"].ToString().Trim();
                    string mascara = lector["mascara"].ToString().Trim(); 

                    Q = "select id_bitacora,Convert(char(10),fecha,103) as fecha,quien,que,aquien,intentos from SICOP_BITACORA ";
                    Q += " where Isnull(fh_envio_bp,'')='' and centralizado='True' and id_agencia='" + id_agencia.Trim() + "'";                    
                    Q += " and Isnull(intentos,0) < " + this.TotalIntentos;
                    Q += " order by id_bitacora";

                    DataSet dsbit = this.objDB.Consulta(Q);
                    foreach (DataRow xregistrar in dsbit.Tables[0].Rows)
                    {
                        Q = "Select ruta_ejecutable_BPro From SICOPCONFIGXMAQUINA where activo='True' and numero_sucursal='" + id_agencia.Trim() + "'";
                        string RutaEjecutableBPro = this.objDB.ConsultaUnSoloCampo(Q).Trim();
                        if (RutaEjecutableBPro.Trim() != "")
                        {
                            string DirectorioArchivosSICOP = this.objDB.ConsultaUnSoloCampo("Select carpeta_local_ventas From SICOPCONFIGXMAQUINA where activo='True' and numero_sucursal='" + id_agencia.Trim() + "'").Trim();
                            string fecharegistrar = xregistrar["fecha"].ToString().Trim();
                            const string quote = "\"";
                            string res = "";
                            string Sicop = "SICOP";
                            string Comando = quote + RutaEjecutableBPro.Trim() + quote + " {0} {1} {2} {3} {4} {5} {6}"; //this.RutaEjecutableBPro.Trim() + " {0} {1} {2} {3} {4} {5} {6}";   //@""" + this.RutaEjecutableBPro.Trim() + @""" + " {0} {1} {2} {3} {4} {5} {6}";
                            string Sentido = "Exporta";

                            string idbitacora = xregistrar["id_bitacora"].ToString().Trim();
                            string CodigoLeido = xregistrar["aquien"].ToString().Trim();

                            Q = " Update SICOP_BITACORA set intentos = Isnull(intentos,0) + 1 where id_bitacora=" + idbitacora.Trim();
                            this.objDB.EjecUnaInstruccion(Q);

                            //primero acutalizamos la fecha en la base de datos correspondiente.
                            //es la fecha de la bitacora                        
                            res = RegistrarSalida(CodigoLeido.Trim(), fecharegistrar.Trim(), id_agencia.Trim());
                            if (res.IndexOf("Error:") == -1)
                            { //No hubo error
                                try
                                {
                                    //"C:\Users\omorales\Desktop\Business Pro SICOP.exe" SICOP GMI GAZM_ZARAGOZA Exporta C:\SiCoP\Generar\ SICOP_PROSPECTOS_TEMP_DMS.TXT 3N1CK3CD9DL259265 1000 25832
                                    //"C:\Users\omorales\Desktop\Business Pro SICOP.exe" SICOP GMI GAZM_ZARAGOZA Exporta C:\SiCoP\Generar\ SICOP_PROSPECTOS_TEMP_DMS.TXT 3N1CK3CD9DL259265
                                    Comando = string.Format(Comando, Sicop, UsuarioBPRo1, BDBPRo1, Sentido, DirectorioArchivosSICOP.Trim(), "parametro_ocioso.txt", CodigoLeido.Trim());
                                    LanzaEjecucion(Comando); //lo deja en una sola carpeta.                                 
                                    Utilerias.WriteToLog("Se ejecutó: " + Comando, "ProcesaBitacora", Application.StartupPath + "\\Log.txt");
                                    //Esperamos un minuto para que le de tiempo a la interfaz a crear el archivo.
                                    Thread.Sleep(Convert.ToInt16(this.MinutosEsperaraBPro) * 60000); //20150505 En lugar del fsw_created.
                                    procesaArchivoGeneradoporBPro(carpeta_local_ventas.Trim() + "\\" + mascara.Trim(), mascara.Trim());
                                    Utilerias.WriteToLog("", "", Application.StartupPath + "\\Log.txt");
                                    Thread.Sleep(20000);
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine(ex.Message);
                                    Utilerias.WriteToLog(ex.Message, "ProcesaBitacora", Application.StartupPath + "\\Log.txt");
                                }
                            }
                            else
                            {
                                Utilerias.WriteToLog("No se actualizo la fecha de salida en BPro: " + res + " id_bitacora: " + idbitacora + " vin: " + CodigoLeido, "ProcesaBitacora", Application.StartupPath + "\\Log.txt");
                            }
                        }//de si existe ruta del ejecutable de BPro.
                        else {
                            Utilerias.WriteToLog(" No se encontró ruta de ejecutable de BPro para la agencia: " + id_agencia, "ProcesaBitacora", Application.StartupPath + "\\Log.txt");  
                        }
                    } //del for de cada bitacora por mandar a ejecutar 
                }//del try
                catch (Exception ex)
                {
                    Utilerias.WriteToLog(ex.Message, "ProcesaBitacora", Application.StartupPath + "\\Log.txt");
                    Debug.WriteLine(ex.Message); 
                }
            } //del ciclo sobre cada lector.      
        }

        private void RevisaLectores()
        {
            string Q = "Select * from SICOPCONFIGXMAQUINA where activo='True' order by Convert(int,numero_sucursal)";

            DataSet ds = this.objDB.Consulta(Q);
            foreach (DataRow lector in ds.Tables[0].Rows)
            {
                try
                {
                    string strUsrRemoto = lector["usr_local"].ToString().Trim(); 
                    string strDominio = "";
                    string strIPFileStorage = lector["ip_local"].ToString().Trim();
                    string strCarpetaLocalVins = lector["carpeta_local_vins"].ToString().Trim();
                    string strCarpetaServerDejar = lector["carpeta_server_dejar"].ToString().Trim();
                    string idmaquina = lector["id_maquina"].ToString().Trim();
                    string nombremaquina = lector["nombre"].ToString().Trim();
                    string id_agencia = lector["numero_sucursal"].ToString().Trim();
                    string PassLocal = lector["passw_local"].ToString().Trim();

                    if (strUsrRemoto.IndexOf("\\") > -1)
                    {   // DANDRADE\sistemas     DANDRADE = dominio sistemas=usuario
                        strDominio = strUsrRemoto.Substring(0, strUsrRemoto.IndexOf("\\"));
                        strUsrRemoto = strUsrRemoto.Substring(strUsrRemoto.IndexOf("\\") + 1);
                    }

                    #region funciones de logueo
                    IntPtr token = IntPtr.Zero;
                    IntPtr dupToken = IntPtr.Zero;
                    //primero intentamos el logueo en el servidor remoto
                    bool isSuccess = false;
                    if (strDominio.Trim() != "") //cuando la impersonizacion es en un servidor que pertenece a un dominio entonces es necesario autenticarse haciendo uso del dominio y no de la ip.
                        isSuccess = LogonUser(strUsrRemoto, strDominio, PassLocal.Trim(), LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref token);
                    else
                        isSuccess = LogonUser(strUsrRemoto, strIPFileStorage, PassLocal.Trim(), LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref token);

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

                            FileInfo[] Archivos = new DirectoryInfo(strCarpetaLocalVins).GetFiles("*_*.txt");

                            foreach (FileInfo Archivo in Archivos)
                            {

                                string fechaenarchivo = Archivo.Name.Substring(0, 14);
                                string vinenarchivo = Archivo.Name.Substring(Archivo.Name.IndexOf("_") + 1);
                                vinenarchivo = vinenarchivo.Replace(".txt", "");
                                vinenarchivo = vinenarchivo.Replace(".TXT", "");
                                //convertimos de yyyyMMddHHmmss a  2015-02-13 10:19:21 
                                fechaenarchivo = fechaenarchivo.Substring(0, 4) + "-" + fechaenarchivo.Substring(4, 2) + "-" + fechaenarchivo.Substring(6, 2) + " " + fechaenarchivo.Substring(8, 2) + ":" + fechaenarchivo.Substring(10, 2) + ":" + fechaenarchivo.Substring(12, 2);          

                                string nuevaruta = strCarpetaServerDejar + "\\" + Archivo.Name;
                                if (File.Exists(nuevaruta))
                                    File.Delete(nuevaruta);
                                Archivo.MoveTo(nuevaruta); //Esta linea es la que envia el archivo
                                string quien = "id_maquina=" + idmaquina + " " + nombremaquina + " " + strIPFileStorage.Trim();   //id_maquina=1 ZARAGOZA50 192.168.7.50
                                Utilerias.WriteToLog("Se Recuperó el archivo: " + nuevaruta + " " + quien, "RevisaLectores", Application.StartupPath + "\\Log.txt");
                                //registramos en BD para su posterior procesamiento.
                                Q = " Insert into SICOP_BITACORA (fecha,quien,que,aquien,id_agencia,centralizado)";
                                Q += " values (Convert(datetime,'" + fechaenarchivo + "',121),'" + quien.Trim() + "','Actualizacion exitosa! Id Prospecto:','" + vinenarchivo.Trim() + "','" + id_agencia + "','True')";
                                this.objDB.EjecUnaInstruccion(Q);
                                Utilerias.WriteToLog("Se Registró: " + Q, "RevisaLectores", Application.StartupPath + "\\Log.txt");
                            }//del forech de cada archivo                                                                                   
                        }
                        catch (Exception exe)
                        {
                            Debug.WriteLine(exe.Message);
                            //Utilerias.WriteToLog("FAILURE: \r" + exe.Message + "\r", "RevisaLectores", Application.StartupPath + "\\Log.txt");
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
                    Utilerias.WriteToLog("Error al loguearse  en la carpeta remota \n\r" + ex.Message, "RevisaLectores", Application.StartupPath + "\\Log.txt");
                    Debug.WriteLine(ex.Message);
                }
            
            }//del ciclo de cada lector
        }


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
                this.timerThread.Interval=(Convert.ToInt16(this.Latencia) * 60000);
                this.timerThread.Enabled = true;
                this.timerThread.Start();
                
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
        public string RegistrarSalida(string vin,string fecharegistrar, string id_agencia)
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
                    SqlConnection conBP = new SqlConnection();
                    SqlCommand bp_comand = new SqlCommand();
                    
                    //conociendo el id_agencia procedemos a consultar los datos de conexion en la tabla transferencia
                    Q = "Select ip,usr_bd,pass_bd,nombre_bd,bd_alterna, dir_remoto_xml, dir_remoto_pdf,usr_remoto,pass_remoto, ip_almacen_archivos, smtpserverhost, smtpport, usrcredential, usrpassword ";
                    Q += " From SICOP_TRASMISION where id_agencia='" + id_agencia + "'";

                    DataSet ds = this.objDB.Consulta(Q);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        DataRow regConexion = ds.Tables[0].Rows[0];
                        string strconexionABussinesPro = string.Format("Data Source={0};Initial Catalog={1}; Persist Security Info=True; User ID={2};Password={3}", regConexion["ip"].ToString(), regConexion["nombre_bd"].ToString(), regConexion["usr_bd"].ToString(), regConexion["pass_bd"].ToString());
                 
                        if (conBP.State.ToString().ToUpper().Trim() == "CLOSED")
                        {
                            try
                            {
                                conBP.ConnectionString = strconexionABussinesPro;
                                conBP.Open();
                                bp_comand.Connection = conBP;

                                //teniendo la conexion con la base de datos vamos a actualizar

                                //Q = "update UNI_PEDIUNI set PEN_FECHAENTREGA = Convert(char(10),getdate(),103) where pen_idpedi ='" + id_pedido.Trim()  + "' and pen_numserie ='" + vin.Trim() +  "'";
                                //Q = "update UNI_PEDIUNI set PEN_FECHAENTREGA_REAL = Convert(char(10),getdate(),103) where pen_idpedi ='" + id_pedido.Trim() + "' and pen_numserie ='" + vin.Trim() + "'";
                                //Q = "update UNI_PEDIUNI set PEN_FECHAENTREGA_REAL = Convert(char(10),getdate(),103) ";
                                Q = " update UNI_PEDIUNI set PEN_FECHAENTREGA_REAL = '" + fecharegistrar + "'";
                                Q += " where pen_numserie ='" + vin.Trim() + "'";
                                Q += " and Isnull(PEN_FECHAENTREGA_REAL,'01/01/1900')='01/01/1900'"; //para que solo lo haga una sola vez.
                                bp_comand.CommandText = Q.Trim();
                                
                                int totreg = bp_comand.ExecuteNonQuery();                                

                                if (totreg == 1)
                                {
                                    res = "La salida de la unidad con vin: " + vin.Trim() + " ha quedado registrada.";
                                }
                                else {
                                    res = "La fecha de salida ya fue capturada: " + vin.Trim();
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


        public void procesaArchivoGeneradoporBPro(string ArchivoGenerado, string SoloNombre)
        {       
            if (File.Exists(ArchivoGenerado))
            {         
                //El código escaneado debe ser leido del que contine el archivo txt, para que no se confunda cuando se escanea más de un código.
                string vinenarchivo = ObtenDeArchivo(ArchivoGenerado, "No_Vin");
                if (vinenarchivo.Trim() != "")
                {

                    string ArchivoRenombrado = SoloNombre.Trim();
                    ArchivoRenombrado = ArchivoRenombrado.ToUpper();
                    ArchivoRenombrado = ArchivoRenombrado.Replace(".TXT", "") + "_" + vinenarchivo.Trim() + ".TXT";

                    FileInfo fi = new FileInfo(ArchivoGenerado);

                    ArchivoRenombrado = fi.DirectoryName.Trim() + "\\" + ArchivoRenombrado.Trim();
                    fi.MoveTo(ArchivoRenombrado); // le cambiamos el nombre agregandole al nombre del archivo el _vin.

                    string idprospenarchivo = ObtenDeArchivo(ArchivoRenombrado, "IdProspecto");
                    string tipoventa = ObtenDeArchivo(ArchivoRenombrado, "TipoVenta");
                    string id_maquina = "";
                    string id_bitacora = "";
                    string Q = "Select top 1 * from SICOP_BITACORA where aquien='" + vinenarchivo + "' order by id_bitacora desc ";
                    DataSet ds1 = this.objDB.Consulta(Q);
                    foreach (DataRow regbitacora in ds1.Tables[0].Rows)
                    {
                        id_maquina = regbitacora["quien"].ToString().Trim();
                        id_maquina = id_maquina.Substring(0, id_maquina.IndexOf(" ")).Trim();
                        id_maquina = Utilerias.A_Numero(id_maquina);
                        id_bitacora = regbitacora["id_bitacora"].ToString().Trim();
                    }

                    //actualizamos si es que trae id_sicop: 
                    Q = "Update SICOP_BITACORA set que = 'Actualizacion exitosa! Id Prospecto: " + idprospenarchivo.Trim() + "' , fh_envio_bp = getdate() where id_bitacora=" + id_bitacora.Trim() + " and centralizado = 'True'";
                    this.objDB.EjecUnaInstruccion(Q);

                    try
                    {
                        Q = "Select * from SICOPCONFIGXMAQUINA where id_maquina=" + id_maquina.Trim();
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
                                string rutareal = ConsultaCarpetaDestino(idprospenarchivo.Trim());
                                if (tipoventa == "INTERCAMBIOS")
                                {
                                    rutareal = Application.StartupPath + "\\Procesados\\" + tipoventa.Trim();
                                }

                                nuevaruta = rutareal.Trim() == "" ? nuevaruta.Trim() : rutareal.Trim() + "\\" + ArchivoRenombrado.Trim();

                                if (File.Exists(nuevaruta))
                                    File.Delete(nuevaruta);

                                Archivo.CopyTo(nuevaruta); //Esta linea es la que envia el archivo a la carpeta de SICOP.
                                Utilerias.WriteToLog("Se Envió el archivo a la carpeta de SICOP : " + nuevaruta, "fsw_Created", Application.StartupPath + "\\Log.txt");

                                if (File.Exists(Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim()))
                                    File.Delete(Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim());

                                Archivo.CopyTo(Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim());

                                //enviar correo adjuntando el archivo txt. Si trae idprospecto se envia a Desarrollo, si no trae se envia a Operacion                                    
                                EnviaCorreo(vinenarchivo.Trim(), idprospenarchivo, Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim(), "fswCreated", id_maquina, NumeroSucursal, strIPMaquina, strNombreMaquina, strEnviar, tipoventa.Trim());
                                if (File.Exists(Archivo.FullName))
                                {
                                    //Utilerias.WriteToLog("Se borra el archivo: " + Archivo.FullName, "fsw_Created", Application.StartupPath + "\\Log.txt");
                                    Archivo.Delete();
                                }
                            }
                            catch (Exception exe)
                            {
                                Debug.WriteLine(exe.Message);
                                Utilerias.WriteToLog("Error: \r" + exe.Message + "\r", "fsw_Created", Application.StartupPath + "\\Log.txt");
                            }

                            #region Proceso Antiguo de transferencia del archivo a otro servidor
                            /*

                            #region funciones de logueo
                            IntPtr token = IntPtr.Zero;
                            IntPtr dupToken = IntPtr.Zero;
                            //primero intentamos el logueo en el servidor remoto
                            bool isSuccess = false;
                            if (strDominio.Trim() != "") //cuando la impersonizacion es en un servidor que pertenece a un dominio entonces es necesario autenticarse haciendo uso del dominio y no de la ip.
                                isSuccess = LogonUser(strUsrRemoto, strDominio, Pass.Trim(), LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref token);
                            else
                                isSuccess = LogonUser(strUsrRemoto, strIPFileStorage, Pass.Trim(), LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref token);

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
                                    FileInfo Archivo = new FileInfo(e.FullPath);
                                    string nuevaruta = CarpetaRemota + "\\" + ArchivoRenombrado.Trim();
                                    if (File.Exists(nuevaruta))
                                        File.Delete(nuevaruta);
                                    Archivo.CopyTo(nuevaruta); //Esta linea es la que envia el archivo

                                    if (File.Exists(Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim()))
                                        File.Delete(Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim());

                                    Archivo.MoveTo(Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim());
                                    Utilerias.WriteToLog("Se Envió el archivo: " + nuevaruta, "fsw_Created", Application.StartupPath + "\\Log.txt");
                                    //enviar correo adjuntando el archivo txt. Si trae idprospecto se envia a Desarrollo, si no trae se envia a Operacion                                    
                                    EnviaCorreo(vinenarchivo.Trim(), idprospenarchivo, Application.StartupPath + "\\Procesados\\" + ArchivoRenombrado.Trim(), "fswCreated", id_maquina, NumeroSucursal,strIPMaquina,strNombreMaquina,strEnviar );
                                }
                                catch (Exception exe)
                                {
                                    Debug.WriteLine(exe.Message);
                                    Utilerias.WriteToLog("FAILURE: \r" + exe.Message + "\r", "fsw_Created", Application.StartupPath + "\\Log.txt");
                                }

                                isSuccess = CloseHandle(token);
                                if (!isSuccess)
                                {
                                    RaiseLastError();
                                }

                            }
                             */
                            #endregion
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
              Utilerias.WriteToLog("El archivo: " + ArchivoGenerado + "no existe", "procesaArchivoGeneradoporBPro", Application.StartupPath + "\\Log.txt");
            }
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
            procesaArchivoGeneradoporBPro(e.FullPath,e.Name);           
        }

        public string EnviaCorreo(string vin, string idsicop, string rutaarchivocreado, string desde, string Id_Maquina, string NumeroSucursal, string strIPMaquina,string NombreMaquina,string Enviar,string TipoVenta)
        {
            string res = "";
            try
            {
                if (Enviar.Trim() == "True")
                {
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
                        if (idsicop.Trim() == "")
                            mensaje.Subject = "Interfaz SICOP vs BPro aviso automático V.I.N. : " + vin.Trim() + " Sin id prospecto de SICOP ";
                        else
                            mensaje.Subject = "Interfaz SICOP vs BPro aviso automático V.I.N. : " + vin.Trim() + " Id Prospecto: " + idsicop.Trim();

                        string Remitente = "Sistemas de Grupo Andrade";

                        //Si trae idprospecto se envia a Desarrollo, si no trae se envia a Operacion
                        string campoconsultar = idsicop == "" ? "interesados_operacion" : "interesados_desarrollo";
                        Q = "Select " + campoconsultar + " from SICOPCONFIGXMAQUINA where id_maquina=" + Id_Maquina;

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
                            cadenaLog += " INTERFAZ CENTRALIZADA " + "\n" + "\r";
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
            this.ntiBalloon.ShowBalloonTip(1, "SICOP CODIGOS DE BARRAS SALIDAS", " Revisando Lectores ", ToolTipIcon.Info);
               RevisaLectores();
            this.ntiBalloon.ShowBalloonTip(1, "SICOP CODIGOS DE BARRAS SALIDAS", " Procesando Bitacora ", ToolTipIcon.Info);
               ProcesaBitacora();
        }

        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            this.ntiBalloon.Icon = this.Icon;
            this.ntiBalloon.Text = "SICOP CODIGOS DE BARRAS SALIDAS";
            this.ntiBalloon.Visible = true;
            this.ntiBalloon.ShowBalloonTip(1, "SICOP CODIGOS DE BARRAS SALIDAS", " En espera de instrucciones ", ToolTipIcon.Info);
            this.Hide();
            this.Visible = false;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.ntiBalloon.Visible = false;
            this.ntiBalloon = null;
        }

        
        
    }
}
