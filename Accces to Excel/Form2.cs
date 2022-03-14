/*
 * Creado por SharpDevelop.
 * Usuario: Lanze Zager
 * Fecha: 09/06/2015
 * Hora: 03:17 p. m.
 * 
 * Para cambiar esta plantilla use Herramientas | Opciones | Codificación | Editar Encabezados Estándar
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Accces_to_Excel
{
	/// <summary>
	/// Description of Form2.
	/// </summary>
	public partial class Form2 : Form
	{
		public Form2()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		void Form2Load(object sender, EventArgs e)
		{
			this.Top = (Screen.PrimaryScreen.WorkingArea.Height - this.Height) / 2;
			this.Left = (Screen.PrimaryScreen.WorkingArea.Width - this.Width) / 2;		
		}
		

		
		void Button1Click(object sender, EventArgs e)
		{
			this.Close();			
		}
		
		void Label4Click(object sender, EventArgs e)
		{
			
		}
	}
}
