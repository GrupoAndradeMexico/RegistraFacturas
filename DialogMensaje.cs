using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace RegistraFactura
{
    public partial class DialogMensaje : Form
    {
        public static bool Acepto = false;
        string TituloForma = "";
        string MensajeMostrar = "";
        bool EsperarRespuesta = false;
        int SegundosMostrar = 0;
        string TextoLetreroEspera = "Por favor espere";
        int NumPuntosMostrar = 10;
        bool MostrarDegradado = false;
        Color colorfondoforma = Color.White;
        Color colortextomensaje = Color.Black;
        string imagenicono = String.Empty;
        string[] EtiquetasBotones;
        int TamanoGrandeMensaje = 60; // si el mensaje a mostrar es mayor a esta cantidad entonces recoloca y agranda los objetos para poder mostrarlo
        int NumRenglones = 3; //si el mensaje tiene entre MAS de este número de renglones entonces lo recoloca y agranda

        public DialogMensaje(string TituloForma,string MensajeMostrar,bool EsperarRespuesta,string TextoEspera,int SegundosMostrar,bool MostrarDegradado,Color Colorfondofrm, Color colortextomensaje,string imagenicono,string EtiquetasBotones)
        {
            InitializeComponent();

            this.TituloForma = TituloForma.Trim();
            this.MensajeMostrar = MensajeMostrar.Trim();
            this.EsperarRespuesta = EsperarRespuesta;
            this.SegundosMostrar = SegundosMostrar;
            this.TextoLetreroEspera = TextoEspera;
            this.MostrarDegradado = MostrarDegradado;
            this.colorfondoforma = Colorfondofrm;
            this.colortextomensaje = colortextomensaje;
            this.imagenicono = imagenicono;
            this.EtiquetasBotones = EtiquetasBotones.Split(',');

            
        }

        private void DialogMensaje_Load(object sender, EventArgs e)
        {
            this.Text = this.TituloForma;
            this.lblMensaje.Text = this.MensajeMostrar;
            this.letrero.Text = this.TextoLetreroEspera;
            this.BackColor = this.colorfondoforma;
            this.lblMensaje.ForeColor = this.colortextomensaje;

            string[] Arr = this.MensajeMostrar.Split('\n');
            if (this.MensajeMostrar.Length > this.TamanoGrandeMensaje || Arr.Length > this.NumRenglones)
            {
                AgrandayRecoloca();
            }

            if (this.imagenicono != "")
            {
                ColocaIcono();
            }

            if (this.MostrarDegradado == false)
            {
                this.Opacity = 100;
            }

            if (EsperarRespuesta == false)
            {
                this.letrero.Visible = true;
                //timer1 es el encargado de disparar el cierre automático de la forma
                timer1.Interval = SegundosMostrar * 1000;
                timer1.Enabled = true;
                timer1.Start();
                this.Size = new Size(626,200);
                this.ControlBox = true;
            }

            //tick es el encargado de simular el avance en los puntos del texto de espera y aumentar o disminuir la opacidad de la forma.
            tick.Start();
            tick.Tick += new EventHandler(tick_Tick);

            if (EsperarRespuesta == true)
            {   int cont=0;
                foreach (string etiq in this.EtiquetasBotones)
                {
                    switch (cont)
                    {
                        case 0: this.cmdCancelar.Text = etiq.Trim();
                                this.cmdCancelar.Visible = true;
                            break;
                        case 1: this.cmdAceptar.Text = etiq.Trim();
                                this.cmdAceptar.Visible = true;
                            break;
                    }
                    cont++;
                }                                 
            }            
        }

        private bool ColocaIcono()
        {
            bool res = false;
          switch (this.imagenicono.ToUpper().Trim())
          {
              case "ERROR":
                  if (File.Exists(Application.StartupPath + @"\Error.jpg")){
                      Bitmap imagen = new Bitmap(Application.StartupPath + @"\Error.jpg");
                      this.pBIcono.Image = (Image)imagen;                      
                      this.pBIcono.Visible = true;
                      res = true;
                  }
                  break;
              case "INTERROGACION":
                  if (File.Exists(Application.StartupPath + @"\Interrogacion.jpg"))
                  {
                      Bitmap imagen = new Bitmap(Application.StartupPath + @"\Interrogacion.jpg");
                      this.pBIcono.Image = (Image)imagen;
                      this.pBIcono.Visible = true;
                      res = true;
                  }
                  break;
              case "PALOMA":
                  if (File.Exists(Application.StartupPath + @"\Paloma.jpg"))
                  {
                      Bitmap imagen = new Bitmap(Application.StartupPath + @"\Paloma.jpg");
                      this.pBIcono.Image = (Image)imagen;
                      this.pBIcono.Visible = true;
                      res = true;
                  }
                  break;
              case "INFORMACION":
                  if (File.Exists(Application.StartupPath + @"\Informacion.jpg"))
                  {
                      Bitmap imagen = new Bitmap(Application.StartupPath + @"\Informacion.jpg");
                      this.pBIcono.Image = (Image)imagen;
                      this.pBIcono.Visible = true;
                      res = true;
                  }
                  break;
              case "MSN":
                  if (File.Exists(Application.StartupPath + @"\msn.jpg"))
                  {
                      Bitmap imagen = new Bitmap(Application.StartupPath + @"\msn.jpg");
                      this.pBIcono.Image = (Image)imagen;
                      this.pBIcono.Visible = true;
                      res = true;
                  }
                  break;                
          }
          return res;
        }

        void tick_Tick(object sender, EventArgs e)
        {
            if (this.EsperarRespuesta == false)
            {
                if (letrero.Text.Length < (this.TextoLetreroEspera.Length + this.NumPuntosMostrar))
                    letrero.Text += ".";
                else
                    letrero.Text = this.TextoLetreroEspera.Trim();
            }
            if (this.Opacity < 100 && this.MostrarDegradado==true)
                this.Opacity = this.Opacity + 0.10;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.timer1.Stop();
            this.tick.Stop();
            this.Close();
        }

        private void cmdAceptar_Click(object sender, EventArgs e)
        {
            Acepto = true;
            this.timer1.Stop();
            this.tick.Stop(); 
            this.Close();
        }

        private void cmdCancelar_Click(object sender, EventArgs e)
        {
            Acepto = false;
            this.timer1.Stop();
            this.tick.Stop(); 
            this.Close();
        }

        /// <summary>
        /// Agranda la forma y recoloca los botones en el caso de que sea un mensaje muy grande por mostrar
        /// </summary>
        private void AgrandayRecoloca()
        {
            this.Size = new Size(801, 327);
            this.letrero.Location = new Point(7,202); 
            this.cmdAceptar.Location = new Point(71,230);
            this.cmdCancelar.Location= new Point(360,230) ;
            this.lblMensaje.Size = new Size(622, 190);                        
        }

        private void DialogMensaje_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.tick.Stop();
            this.timer1.Stop(); 
        }

    }
}