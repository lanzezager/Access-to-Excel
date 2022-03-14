/*
 * Creado por SharpDevelop.
 * Usuario: Lanze Zager
 * Fecha: 22/09/2015
 * Hora: 03:34 p. m.
 * 
 * Para cambiar esta plantilla use Herramientas | Opciones | Codificación | Editar Encabezados Estándar
 */
using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;

namespace Accces_to_Excel
{
	/// <summary>
	/// Description of Mensaje.
	/// </summary>
	public partial class Mensaje : Form
	{
		public Mensaje(String errorz, String lineaz)
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			this.linea = lineaz;
			this.error = errorz;
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		String error,linea;
		void Button1Click(object sender, EventArgs e)
		{
			this.Dispose();
		}
		
		void MensajeLoad(object sender, EventArgs e)
		{
			textBox1.Clear();
			textBox2.Clear();
			textBox1.Text = error;
			textBox2.Text = linea;
		}
	}
}
