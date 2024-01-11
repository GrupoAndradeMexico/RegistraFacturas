using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Collections.Specialized;
using System.Data;
using ICSharpCode.SharpZipLib.Zip;

namespace RegistraSalida
{
    class Utilerias
    {
        //ConexionBD objDB = new ConexionBD("BDLOCAL"); 
        
        public Utilerias()
        { 
          
        }

        public void EmiteSonidoLlamado(string Directorio)
        {
            string nomArch = Directorio + "\\LlamarTurno.wav";

            if (File.Exists(nomArch))
            {
                string extension = Path.GetExtension(nomArch);
                string sinExt = Path.GetFileNameWithoutExtension(nomArch);
                if (extension.ToUpper() == ".WAV")
                {
                    try
                    {
                        System.Media.SoundPlayer soundPlayer = new System.Media.SoundPlayer();
                        soundPlayer.SoundLocation = nomArch;
                        soundPlayer.Play();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                }
            }
        }


        public void EmiteSonidoAleatorio(string Directorio)
        {
            //cargamos todos los sonidos disponibles (sonidos wav)
                string[] Archivos = Directory.GetFiles(Directorio, "*.wav");
                //ahora elegimos uno al azar es decir aleatoriamente
                Random numeroaleatorio = new Random();
                int num = numeroaleatorio.Next(0, Archivos.Length);
                //el nombre del archivo ahora es:
                string nomArch = Archivos[num].ToString();

                //nomArch = this.InitialDirectory + "PIG.WAV"; 
                string extension = Path.GetExtension(nomArch);
                string sinExt = Path.GetFileNameWithoutExtension(nomArch);
                if (extension.ToUpper()==".WAV") 
                {
                    try
                    {
                        System.Media.SoundPlayer soundPlayer = new System.Media.SoundPlayer();
                        soundPlayer.SoundLocation = nomArch;
                        soundPlayer.Play();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message); 
                    }
                }
        }

        
        //Dada un diccionario y un valor regresa la llave si la encuentra
        //si no regresa la cadena vacia
        public string DameLLave(StringDictionary coleccion, string valor)
        {
            string res = "";
            foreach (string llave in coleccion.Keys)
            {
                string valorcol = coleccion[llave].Trim();
                if (valorcol.Equals(valor))
                {
                    res = llave;
                    return res;
                }
            }
            return res;
        }

        public static string DameLlave(NameValueCollection coleccion, string valor)
        {
            string res = "";
            foreach (string llave in coleccion.Keys)
            {
                string valorcol = coleccion[llave].Trim();
                if (valorcol.Equals(valor))
                {
                    res = llave;
                    return res;
                }
            }
            return res;
        }

        /// <summary>
        ///"lee toda la cadena número que se le pasa como  argumento y se queda únicamente con los caracteres que son dígitos - o ." 
        /// </summary>
        public static string A_Numero(string numero)
        {
            string res = numero.Trim();
            numero = numero.Trim();
            if (numero.Length > 0)
            {
                //char[] cant = new char[numero.Length]; 
                int i = 0;
                int numdigitos = 0;
                string digito = "";
                string cantidad = "";
                string digitos = "-0123456789.,";
                while (i < numero.Length)
                {
                    digito = numero.Substring(i, 1);
                    if (digitos.IndexOf(digito) != -1)
                    {
                        cantidad += numero[i].ToString();
                        numdigitos++;
                    }
                    i++;
                }
                //ya tenemos en cantidad el número con sólo dígitos 
                //ahora hacemos la conversion
                res = cantidad.Trim();
            }
            return res;
        }

        //regresa true si lo que recibe no contiene letras o caracteres que no pertenezcan al formato #,.
        //donde el # es cualquier cantidad de dígitos.
        public static bool EsNumero(string numero)
        {
            bool res = true;
            int i = 0;
            string digito = "";
            string digitos = "-0123456789.,";
            while (i < numero.Length && res == true)
            {
                digito = numero.Substring(i, 1);
                if (i == 0 && digito == "-" && numero.Length < 2) //solo es el signo - , no un número negativo
                    res = false;
                if (i == 0 && digito == "," && numero.Length < 2) //solo es el signo ',' , no un número
                    res = false;
                if (i == 0 && digito == "." && numero.Length < 2) //solo es el signo '.' , no un número
                    res = false;

                if (res != false)
                {
                    if (digitos.IndexOf(digito) == -1) //en cuanto no esta el caracter entre los dígitos devuelve falso
                    {
                        res = false;
                    }
                }
                i++;
            }
            return res;
        }



        public static string A_FechaU(string fecha, string formato)
        {
            DateTime fechaaux;
            try
            {
                if (fecha.Length == 6)
                { //al año le faltan 2 dígitos esperamos que el formato sea yyMMdd
                    string anioaux = fecha.Substring(0, 2);
                    int anio = A_Entero(anioaux);
                    if (anio >= 20 && anio <= 99) //la fecha esta en el siglo pasado
                    {
                        anio = 1900 + anio;
                    }
                    else if (anio >= 0 && anio < 20) //LJBA este codigo solo servirá hasta el 2019
                    {
                        anio = 2000 + anio;
                    }
                    fecha = anio.ToString() + "/" + fecha.Substring(2, 2) + "/" + fecha.Substring(4, 2);
                }
                if (fecha.Length == 8)
                {
                    //'LJBA 20101014 vericacion del formato de fecha
                    string ls_aux_fmto_fecha_yyyymmdd = fecha.Substring(4, 4) + fecha.Substring(2, 2) + fecha.Substring(0, 2);
                    string ls_aux_fmto_fecha_ddmmyyyy = fecha.Substring(0, 2) + fecha.Substring(2, 2) + fecha.Substring(4, 4);
                    string ls_aux_fecha = "";
                    Double ld_aux_fec_numero = 0;

                    ld_aux_fec_numero = Convert.ToDouble(ls_aux_fmto_fecha_yyyymmdd);
                    if (ld_aux_fec_numero > 19010101 && ld_aux_fec_numero < 20351231)
                    {
                        ls_aux_fmto_fecha_yyyymmdd = fecha.Substring(4, 4) + "/" + fecha.Substring(2, 2) + "/" + fecha.Substring(0, 2);
                        if (EsFecha(ls_aux_fmto_fecha_yyyymmdd))
                        {
                            fecha = fecha.Substring(0, 2) + "/" + fecha.Substring(2, 2) + "/" + fecha.Substring(4, 4); //formato dd/MM/yyyy
                        }
                        else
                        {
                            fecha = "01/01/1900";
                        }
                    }
                    else
                    {
                        ld_aux_fec_numero = Convert.ToDouble(ls_aux_fmto_fecha_ddmmyyyy);
                        if (ld_aux_fec_numero != 1011900)
                        {
                            if (ld_aux_fec_numero > 1011900 && ld_aux_fec_numero < 31122015)
                            {
                                ls_aux_fmto_fecha_ddmmyyyy = fecha.Substring(6, 2) + "/" + fecha.Substring(4, 2) + "/" + fecha.Substring(0, 4); //formato dd/MM/yyyy
                                if (EsFecha(ls_aux_fmto_fecha_ddmmyyyy))
                                {
                                    fecha = ls_aux_fmto_fecha_ddmmyyyy;
                                }
                                else
                                {
                                    fecha = "01/01/1900";
                                }
                            }
                        }
                        else //quiere decir que viene 01011900
                        {
                            fecha = "01/01/1900";  //fecha.Substring(5, 4) + fecha.Substring(3, 2) + fecha.Substring(1, 2);
                        }
                    }
                }//de que tiene 8 caracteres

                fechaaux = Convert.ToDateTime(fecha);
                fecha = fechaaux.ToString(formato);
            }
            catch (Exception e)
            {
                string a = e.Message;
                return fecha;
            }
            return fecha;
        }

        /// <summary>
        /// Regresa la fecha en formato dd/MM/yyyy
        /// </summary>
        /// <param name="fecha">cadena que representa la fecha a formatear</param>
        /// <returns>cadena con la fecha en formato dd/MM/yyyy</returns>
        public static string A_Fecha(string fecha)
        {

            DateTime fechaaux;
            try
            {
                if (fecha.Length == 6)
                { //al año le faltan 2 dígitos esperamos que el formato sea yyMMdd
                    string anioaux = fecha.Substring(0, 2);
                    int anio = A_Entero(anioaux);
                    if (anio >= 20 && anio <= 99) //la fecha esta en el siglo pasado
                    {
                        anio = 1900 + anio;
                    }
                    else if (anio >= 0 && anio < 20) //LJBA este codigo solo servirá hasta el 2019
                    {
                        anio = 2000 + anio;
                    }
                    fecha = anio.ToString() + "/" + fecha.Substring(2, 2) + "/" + fecha.Substring(4, 2);
                }
                if (fecha.Length == 8)
                { 
                                 //'LJBA 20101014 vericacion del formato de fecha
                                string ls_aux_fmto_fecha_yyyymmdd = fecha.Substring(5,4) + fecha.Substring(3, 2) + fecha.Substring(1, 2);
                                string ls_aux_fmto_fecha_ddmmyyyy = fecha.Substring(1, 2) + fecha.Substring(3, 2) + fecha.Substring(5, 4);
                                string ls_aux_fecha="";
                                Double ld_aux_fec_numero=0;
                                
                                ld_aux_fec_numero = Convert.ToDouble(ls_aux_fmto_fecha_yyyymmdd);
                                if (ld_aux_fec_numero > 19010101 && ld_aux_fec_numero < 20351231)
                                {
                                    ls_aux_fmto_fecha_yyyymmdd = fecha.Substring(5, 4) + "/" + fecha.Substring(3, 2) + "/" + fecha.Substring(1, 2);
                                    if (EsFecha(ls_aux_fmto_fecha_yyyymmdd)) 
                                    {
                                      fecha=fecha.Substring(1, 2) + "/" + fecha.Substring(3, 2) + "/" + fecha.Substring(5, 4); //formato dd/MM/yyyy
                                    }
                                    else
                                    {
                                      fecha="01/01/1900";    
                                    }
                                }
                                else
                                {
                                   ld_aux_fec_numero = Convert.ToDouble(ls_aux_fmto_fecha_ddmmyyyy);
                                   if (ld_aux_fec_numero != 1011900) 
                                    {
                                        if (ld_aux_fec_numero > 1011900 && ld_aux_fec_numero < 31122015)
                                        {
                                         ls_aux_fmto_fecha_ddmmyyyy = fecha.Substring(7, 2) + "/" + fecha.Substring(5, 2) + "/" + fecha.Substring(1, 4); //formato dd/MM/yyyy
                                         if (EsFecha(ls_aux_fmto_fecha_ddmmyyyy))
                                         {
                                            fecha=ls_aux_fmto_fecha_ddmmyyyy;                                                        
                                         }
                                         else
                                            {
                                             fecha="01/01/1900";
                                            }
                                        }
                                    }
                                   else //quiere decir que viene 01011900
                                    {
                                        fecha = "01/01/1900";  //fecha.Substring(5, 4) + fecha.Substring(3, 2) + fecha.Substring(1, 2);
                                    }
                                }
                }//de que tiene 8 caracteres

                fechaaux = Convert.ToDateTime(fecha);
                fecha = fechaaux.ToString("dd/MM/yyyy");
            }
            catch (Exception e)
            {
                string a = e.Message;
                return fecha;
            }
            return fecha;
        }

        public static bool EsFecha(string fecha)
        {
            bool res = false;
            DateTime fechaaux;
            if (fecha.IndexOf("/") == -1)
                return res;
            try
            {
                fechaaux = Convert.ToDateTime(fecha);
                res = true;
            }
            catch (Exception e)
            {
                string al = e.Message;
                return false;
            }
            return res;
        }
        /// <summary>
        ///Recibe una cadena que supone tiene inmerso un número p.e. "12r4a"
        ///Regresa el número entero que se encuentra en ella --> 124
        ///si no hay un número entero, entonces regresa un 0 
        /// </summary>
        public static int A_Entero(string numero)
        {
            int res = 0;
            numero = numero.Trim();
            if (numero.Length > 0)
            {
                //char[] cant = new char[numero.Length]; 
                int i = 0;
                int numdigitos = 0;
                string digito = "";
                string cantidad = "";
                string digitos = "-0123456789";
                while (i < numero.Length)
                {
                    digito = numero.Substring(i, 1);
                    if (digitos.IndexOf(digito) != -1)
                    {
                        cantidad += numero[i].ToString();
                        numdigitos++;
                    }
                    i++;
                }
                //ya tenemos en cantidad el número con sólo dígitos 
                //ahora hacemos la conversion
                if (cantidad.Length > 0)
                    res = Convert.ToInt16(cantidad);
            }
            return res;
        }
        /// <summary>
        /// Dado un número regresa solo la parte entera del número
        /// </summary>
        /// <param name="numero">el número que se va a partir</param>
        /// <returns>la parte entera del número dado</returns>
        public static decimal DaParteEntera(double numero)
        {
            //double Fractional = numero - (long)numero;
            return Convert.ToDecimal(numero - (numero - (long)numero));              
        }
        /// <summary>
        /// Dado un número regresa solo la parte decimal i.e. la mantisa
        /// </summary>
        /// <param name="numero">el número que se va a partir</param>
        /// <returns>la parte decimal del número dado</returns>
        public static double DaParteDecimal(double numero)
        {
            return (numero - (long)numero);
            //double Fractional = tiempo.TotalSeconds - (long)tiempo.TotalSeconds;
            //decimal Entera = Convert.ToDecimal(tiempo.TotalSeconds - Fractional); 
        }


        /// <summary>
        /// Dado un Número, lo regresa con el formato de moneda local sin el signo de pesos
        /// </summary>
        /// <param name="Numero"></param>        
        /// <returns></returns>

        public static string FormateaNumero(string Numero)
        {
            try
            {
                double num = Convert.ToDouble(Numero);
                Numero = num.ToString("c");
                Numero = Numero.Replace("$", ""); 
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message); 
            }
            return Numero;
        }



        #region Para implementar El Select Distinct en una DataTable
        private static bool ColumnEqual(object A, object B)
        {

            // Compares two values to see if they are equal. Also compares DBNULL.Value.
            // Note: If your DataTable contains object fields, then you must extend this
            // function to handle them in a meaningful way if you intend to group on them.

            if (A == DBNull.Value && B == DBNull.Value) //  both are DBNull.Value
                return true;
            if (A == DBNull.Value || B == DBNull.Value) //  only one is DBNull.Value
                return false;
            return (A.Equals(B));  // value type standard comparison
        }

        public static DataTable SelectDistinct(string TableName, DataTable SourceTable, string FieldName)
        {
            DataTable dt = new DataTable(TableName);
            dt.Columns.Add(FieldName, SourceTable.Columns[FieldName].DataType);

            object LastValue = null;
            foreach (DataRow dr in SourceTable.Select("", FieldName))
            {
                if (LastValue == null || !(ColumnEqual(LastValue, dr[FieldName])))
                {
                    LastValue = dr[FieldName];
                    dt.Rows.Add(new object[] { LastValue });
                }
            }
            //if (ds != null)
            //    ds.Tables.Add(dt);
            return dt;
        }


        #endregion

        #region Crear el archivo de registro de Articulos por surtir.
        public void RegistraPorSurtir(string id_producto,string nombre,decimal cantidad_actual,string rutaArchivo)
        { //cada registro (línea) en el archivo tiene el lay-out  
          //id_producto - nombre,cantidad_actual,fecharegistro.
          //cuando se pasa el id_pruducto se busca su registro, si ya esta en el archivo, no hace nada.
          //si el registro del producto no se encuentra en el archivo, se agrega.
            try
            {
                FileStream fs;
                if (File.Exists(rutaArchivo))
                {                                                            
                    //antes de escribir en el archivo buscamos por el id del producto.
                    ArrayList arrText = new ArrayList();                    
                    StreamReader sr = new StreamReader(rutaArchivo);
                    string sLine = "";
                    while (sLine != null)
                    { 
                        sLine = sr.ReadLine();                        
                        if (sLine != null)
                        {   //solo agregamos los id de producto
                            sLine = sLine.Substring(0,sLine.IndexOf("-"));
                            arrText.Add(sLine);                     
                        }                            
                    }
                    sr.Close();
                    //si no contiene el id del producto entonces lo agregamos al archivo.
                    if (!arrText.Contains(id_producto))
                    {
                        fs = new FileStream(rutaArchivo, FileMode.Append, FileAccess.Write, FileShare.None);
                        StreamWriter sw = new StreamWriter(fs,Encoding.ASCII);
                        sw.WriteLine(id_producto.Trim() + "-" + nombre.Trim() + "," + cantidad_actual.ToString().Trim() + "," + DateTime.Now.ToString());
                        sw.Close();
                        fs.Close();            
                    }

                }
                else
                {//El archivo no existe por lo tanto lo creamos y escribimos en el 
                    fs = new FileStream(rutaArchivo, FileMode.Create, FileAccess.Write, FileShare.None);
                    StreamWriter sw = new StreamWriter(fs, Encoding.ASCII);
                    sw.WriteLine(id_producto.Trim() + "-" + nombre.Trim() + "," + cantidad_actual.ToString().Trim() + "," + DateTime.Now.ToString());
                    sw.Close();
                    fs.Close();            
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

        }
        #endregion

        #region Escribir en un archivo de texto el log.
        public static bool LimpiaArchivoLog(string rutaArchivo)
        {
            bool res = false;
            FileStream fs = null;
            try
            {
                fs = new FileStream(rutaArchivo, FileMode.Create, FileAccess.Write, FileShare.None);                               
                fs.Close();
                res = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            return res;
        } 


        public static void WriteToLog(string message,string Desde, string rutaArchivo)
        {
            FileStream fs =  null;
            StreamWriter sw = null;
            try
            {
                if (File.Exists(rutaArchivo))
                {
                    fs = new FileStream(rutaArchivo, FileMode.Append, FileAccess.Write, FileShare.None);
                    sw = new StreamWriter(fs, Encoding.UTF8);
                    if (message.Trim() != "")
                        sw.WriteLine(DateTime.Now.ToString() + "\t" + message.Trim() + ", desde: " + Desde);
                    else
                        sw.WriteLine(" ");
                    sw.Close();
                    fs.Close();
                }
                else
                {//El archivo no existe por lo tanto lo creamos y escribimos en el 
                    fs = new FileStream(rutaArchivo, FileMode.Create, FileAccess.Write, FileShare.None);
                    sw = new StreamWriter(fs, Encoding.UTF8);
                    if (message.Trim() != "")
                        sw.WriteLine(DateTime.Now.ToString() + "\t" + message.Trim() + ", desde: " + Desde);
                    else
                        sw.WriteLine(" ");
                    sw.Close();
                    fs.Close();
                }
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
        #endregion


        public static string DescomprimirZip(string sRuta, string ArchZip)
        {
            string res = "";
            try
            {
                FileInfo Archivo = new FileInfo(ArchZip);
                if (Archivo.Extension.ToUpper().Trim() == ".ZIP")
                {
                    FastZip fZip = new FastZip();
                    fZip.ExtractZip(ArchZip, sRuta, "");
                }
                if (Archivo.Extension.ToUpper().Trim() == ".TAR")
                {
                    ICSharpCode.SharpZipLib.Tar.TarArchive ta = ICSharpCode.SharpZipLib.Tar.TarArchive.CreateInputTarArchive(new FileStream(Archivo.FullName, FileMode.Open, FileAccess.Read));
                    //ta.ProgressMessageEvent += delegate(TarArchive tarArchive, TarEntry tarEntry tring msg)
                    //{
                    //   Log(tarEntry.Name + " extracted");
                    //};
                    ta.ExtractContents(sRuta);
                    ta.Close();
                }
                res = sRuta;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                //Utilerias.WriteToLog("Error en Descompresion: " + ex.Message, "DescomprimirZip", Application.StartupPath + "\\Log.txt"); 
            }
            return res;
        }


        //public static string ComprimirZip(string sRuta,string ArchZip,string filtro)
        //    {
        //        string res = "";
        //        bool comprimio = false;
        //        try
        //        {
        //            ZipOutputStream zipOut = new ZipOutputStream(File.Create(ArchZip));
        //            foreach (string fName in Directory.GetFiles(sRuta,filtro))
        //            {
        //                FileInfo fi = new FileInfo(fName);
        //                ZipEntry entry = new ZipEntry(fi.Name);
        //                FileStream sReader = File.OpenRead(fName);
        //                byte[] buff = new byte[Convert.ToInt32(sReader.Length)];
        //                sReader.Read(buff, 0, (int)sReader.Length);
        //                entry.DateTime = fi.LastWriteTime;
        //                entry.Size = sReader.Length;
        //                sReader.Close();
        //                zipOut.PutNextEntry(entry);
        //                zipOut.Write(buff, 0, buff.Length);
        //                comprimio = true;
        //            }
        //            zipOut.Finish();
        //            zipOut.Close();
        //            if (File.Exists(ArchZip) && comprimio)
        //                res = ArchZip.Trim();
        //        }
        //        catch (Exception ex)
        //        {
        //            Debug.WriteLine(ex.Message);
        //        }
        //        return res;
        //    }
    }
}
