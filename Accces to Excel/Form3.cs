/*
 * Creado por SharpDevelop.
 * Usuario: Lanze Zager
 * Fecha: 28/09/2015
 * Hora: 01:50 p. m.
 * 
 * Para cambiar esta plantilla use Herramientas | Opciones | Codificación | Editar Encabezados Estándar
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Accces_to_Excel
{
	/// <summary>
	/// Description of Form3.
	/// </summary>
	public partial class Form3 : Form
	{
		public Form3()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
	String cad_con,combotext,database_nom;	
	DialogResult res;
	
		void Form3Load(object sender, EventArgs e)
		{
			comboBox1.SelectedIndex=0;
		}
		
		void Button2Click(object sender, EventArgs e)
		{
			try{
				
			    res = MessageBox.Show("¿Iniciar Conexión?","INICIAR",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
			    
			    if(res == DialogResult.Yes){
			    	
			    	Puente miPuente1 = this.Owner as MainForm;
			    	miPuente1.conec(cad_con,database_nom);
			    
			    	this.Close();
			    }
			    
			}catch(Exception z)
			{
				MessageBox.Show("Conexión Incorrecta\n"+z);
			}
		}
		
		void Button1Click(object sender, EventArgs e)
		{
			try{
				if(comboBox1.SelectedIndex == 0){
					combotext = "localhost";
				}else{
					combotext = comboBox1.Text;
				}
				Conexion cone = new Conexion();
				cad_con=@"server="+combotext+"; userid="+textBox1.Text+"; password="+textBox2.Text+"; database="+textBox3.Text+"";
				cone.probar(cad_con);
			    MessageBox.Show("La Conexión se estableció Exitosamente");
			    database_nom = textBox3.Text;
			    comboBox1.Enabled = false;
			    textBox1.Enabled = false;
			    textBox2.Enabled = false;
			    textBox3.Enabled = false;
			    button2.Enabled=true;
			}catch(Exception z)
			{
				MessageBox.Show("La Conexión no se pudo realizar");
				button2.Enabled=false;
			}
		}
	}
}
