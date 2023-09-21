using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RegistraFactura
{
    public partial class espera : Form
    {
        public static string mensaje = "";

        public espera(string mensage)
        {
            InitializeComponent();
            mensaje = mensage.Trim(); 
        }

        private void espera_Load(object sender, EventArgs e)
        {
            this.timer1.Enabled = true;
            this.timer1.Start();
            this.timer1.Interval = 1000;
            this.lblMensaje.Text = mensaje;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (this.lblMensaje.Text.Length <= 40)
                this.lblMensaje.Text += ".";
            else
                this.lblMensaje.Text = mensaje.Trim();
        }

        private void espera_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.timer1.Stop();
        }
    }
}
