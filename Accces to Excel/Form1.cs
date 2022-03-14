/*
 * Creado por SharpDevelop.
 * Usuario: Lanze Zager
 * Fecha: 05/06/2015
 * Hora: 11:06 a. m.
 * 
 * Para cambiar esta plantilla use Herramientas | Opciones | Codificación | Editar Encabezados Estándar
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Accces_to_Excel
{
	/// <summary>
	/// Description of Form1.
	/// </summary>
	public partial class Form1 : Form
	{
		public Form1()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//

			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		int tot_cols=0,indice=0,nav_res=1,res_bus=0;
		string[] cols,buscar;
		String txt_buscar,rango_buscar;
		//MainForm principal = new principal.valor_a_buscar(buscar,indice);

		public void receptor(string[] columnas){
			this.tot_cols = columnas.Length;
			this.cols = new string[tot_cols];
			this.cols = columnas;
			
			comboBox1.Items.Add("TODO");
			comboBox1.SelectedItem = "TODO";
			
			for(int j=0;j<tot_cols;j++){
				comboBox1.Items.Add(cols[j].ToString());
			}
		}
		
		public void busqueda(){
			res_bus = 0;
			nav_res = 1;
			txt_buscar = textBox1.Text;
			rango_buscar = comboBox1.SelectedItem.ToString();
			buscar = new string[2];
			buscar[0] = txt_buscar;
			buscar[1] = rango_buscar;
			indice = comboBox1.SelectedIndex;
			//MessageBox.Show("indice: "+indice);
			//Puente bridge = new Puente();
			//bridge.puente_busqueda(buscar,indice);
			Puente miPuente = this.Owner as MainForm;
			miPuente.puente_busqueda(buscar,indice);

		}
		
		public void resultados(int resbus){
			res_bus = resbus;
	
			if(resbus>=1){
				//button2.Enabled=true;
				button3.Enabled=true;
				label3.Text=""+nav_res+" de "+resbus+" coincidencias.";
				label3.Refresh();
			}else{
				button2.Enabled=false;
				button3.Enabled=false;
				label3.Text="0 de 0 coincidencias.";
				label3.Refresh();
				MessageBox.Show("Busqueda sin Resultados", "Access to Excel");
				
			}
		}
		
		public void navegacion(){
		
			if(nav_res <= 1){
				button2.Enabled=false;
			}else{
				button2.Enabled=true;
			}
			
			if(nav_res >= res_bus){
				button3.Enabled=false;
			}else{
				button3.Enabled=true;
			}
			Puente miPuente1 = this.Owner as MainForm;
			miPuente1.puente_res((nav_res-1));
			label3.Text=""+nav_res+" de "+res_bus+" coincidencias.";
			label3.Refresh();
		}
		
		public void reinicio(){
			textBox1.Text="";
			comboBox1.Items.Clear();
			comboBox1.Items.Add("TODO");
			comboBox1.SelectedItem = "TODO";
			tot_cols=0;
			indice=0;
			nav_res=1;
			res_bus=0;
		}
		
		void Form1Load(object sender, EventArgs e)
		{
			this.Top = (Screen.PrimaryScreen.WorkingArea.Height - this.Height) / 2;
			this.Left = (Screen.PrimaryScreen.WorkingArea.Width - this.Width) / 2;
			this.Height=195;
			toolTip1.SetToolTip(this.button2, "Anterior");	
			toolTip1.SetToolTip(this.button3, "Siguiente");	
		}
		
		void Button1Click(object sender, EventArgs e)
		{
			this.Height=245;
			busqueda();
		}
		
		void Button2Click(object sender, EventArgs e)
		{
				nav_res--;
				navegacion();
		}
		
		void Button3Click(object sender, EventArgs e)
		{
				nav_res++;
				navegacion();
		}
		
		void TextBox1KeyPress(object sender, KeyPressEventArgs e)
		{
			if (e.KeyChar == (char)(Keys.Enter))
           {
             this.Height=245;
			 busqueda();
			 this.Focus();
           }
		}
		
		void Form1KeyPress(object sender, KeyPressEventArgs e)
		{
			
		}
		
		void IzquierdaToolStripMenuItemClick(object sender, EventArgs e)
		{
			if (nav_res >1){
				nav_res--;
				navegacion();
			}
		}
		
		void DerechaToolStripMenuItemClick(object sender, EventArgs e)
		{
			if (nav_res < res_bus){
				nav_res++;
				navegacion();
			}
		}
	}
}
