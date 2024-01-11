namespace RegistraSalida
{
    partial class DialogMensaje
    {
        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén utilizando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben eliminar; false en caso contrario, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.lblMensaje = new System.Windows.Forms.Label();
            this.cmdCancelar = new System.Windows.Forms.Button();
            this.cmdAceptar = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.tick = new System.Windows.Forms.Timer(this.components);
            this.letrero = new System.Windows.Forms.Label();
            this.pBIcono = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pBIcono)).BeginInit();
            this.SuspendLayout();
            // 
            // lblMensaje
            // 
            this.lblMensaje.Font = new System.Drawing.Font("Comic Sans MS", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMensaje.Location = new System.Drawing.Point(158, 9);
            this.lblMensaje.Name = "lblMensaje";
            this.lblMensaje.Size = new System.Drawing.Size(435, 116);
            this.lblMensaje.TabIndex = 0;
            this.lblMensaje.Text = "label1";
            // 
            // cmdCancelar
            // 
            this.cmdCancelar.BackColor = System.Drawing.Color.Orange;
            this.cmdCancelar.Cursor = System.Windows.Forms.Cursors.Hand;
            this.cmdCancelar.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdCancelar.ForeColor = System.Drawing.Color.Navy;
            this.cmdCancelar.Location = new System.Drawing.Point(72, 159);
            this.cmdCancelar.Name = "cmdCancelar";
            this.cmdCancelar.Size = new System.Drawing.Size(233, 59);
            this.cmdCancelar.TabIndex = 6;
            this.cmdCancelar.Text = "CANCELAR";
            this.cmdCancelar.UseVisualStyleBackColor = false;
            this.cmdCancelar.Visible = false;
            this.cmdCancelar.Click += new System.EventHandler(this.cmdCancelar_Click);
            // 
            // cmdAceptar
            // 
            this.cmdAceptar.BackColor = System.Drawing.Color.Navy;
            this.cmdAceptar.Cursor = System.Windows.Forms.Cursors.Hand;
            this.cmdAceptar.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdAceptar.ForeColor = System.Drawing.Color.White;
            this.cmdAceptar.Location = new System.Drawing.Point(360, 159);
            this.cmdAceptar.Name = "cmdAceptar";
            this.cmdAceptar.Size = new System.Drawing.Size(233, 59);
            this.cmdAceptar.TabIndex = 5;
            this.cmdAceptar.Text = "ACEPTAR";
            this.cmdAceptar.UseVisualStyleBackColor = false;
            this.cmdAceptar.Visible = false;
            this.cmdAceptar.Click += new System.EventHandler(this.cmdAceptar_Click);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // tick
            // 
            this.tick.Interval = 500;
            // 
            // letrero
            // 
            this.letrero.AutoSize = true;
            this.letrero.BackColor = System.Drawing.Color.Transparent;
            this.letrero.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.letrero.ForeColor = System.Drawing.Color.Navy;
            this.letrero.Location = new System.Drawing.Point(7, 131);
            this.letrero.Name = "letrero";
            this.letrero.Size = new System.Drawing.Size(187, 25);
            this.letrero.TabIndex = 7;
            this.letrero.Text = "Por favor espere";
            this.letrero.Visible = false;
            // 
            // pBIcono
            // 
            this.pBIcono.Location = new System.Drawing.Point(12, 9);
            this.pBIcono.Name = "pBIcono";
            this.pBIcono.Size = new System.Drawing.Size(140, 116);
            this.pBIcono.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pBIcono.TabIndex = 8;
            this.pBIcono.TabStop = false;
            this.pBIcono.Visible = false;
            // 
            // DialogMensaje
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(603, 225);
            this.ControlBox = false;
            this.Controls.Add(this.pBIcono);
            this.Controls.Add(this.letrero);
            this.Controls.Add(this.cmdCancelar);
            this.Controls.Add(this.cmdAceptar);
            this.Controls.Add(this.lblMensaje);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DialogMensaje";
            this.Opacity = 0.1;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Mensaje: ";
            this.Load += new System.EventHandler(this.DialogMensaje_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DialogMensaje_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.pBIcono)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblMensaje;
        private System.Windows.Forms.Button cmdCancelar;
        private System.Windows.Forms.Button cmdAceptar;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Timer tick;
        private System.Windows.Forms.Label letrero;
        private System.Windows.Forms.PictureBox pBIcono;
    }
}