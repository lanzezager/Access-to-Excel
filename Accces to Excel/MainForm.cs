/*
 * Creado por SharpDevelop.
 * Usuario: Lanze Zager
 * Fecha: 21/05/2015
 * Hora: 11:45 a. m.
 * 
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Threading;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;


namespace Accces_to_Excel
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form, Puente
	{
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		//Declaracion de elementos para conexion office
		OleDbConnection conexion = null;
		DataSet dataSet = null;
		OleDbDataAdapter dataAdapter = null;
        DataTable tabla_excel = new DataTable();
       
		
		//Declaracion de Variables
		String cad_con,nom_bd,sql,nom_tabla,tabla,fil_tabla,porcent_text,nom_file,nom_save,reg_pat,reg_pat1,fecha_pago,folio,per,imp_tot,valor_celda,t_d,ext,cad_con_sql,datab_name;
		int filas=0,i=0,j=0,k=0,l=0,m=1,n=0,tot_row=0,tot_col=0,porcent=0,tam_tab=0,camb_td=0,res_bus=0,t_bus=0,x=0,y=0,alcance_bus=0,actual=0;
		double por;
		string[] columnas={"Registro Patronal","Periodo","Tipo DOC","Importe Total","Fecha de Pago","Folio SUA"};
		string[] buscar_params = new string[2];
		string[] busqueda = new string[2];
		string[] fila;
		int[,] cord_res;
		
		SaveFileDialog fichero = new SaveFileDialog();
		Form1 buscar;
		Form3 sqlform;
		
		Conexion consql = new Conexion();
		
		private Thread hilosecundario = null;
		
		public void conectar(){
            //MessageBox.Show(ext);
			if((ext).Equals("mdb")){
					cad_con = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + nom_bd + "';";
					
				}else{
					cad_con = "provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + nom_bd + "';";
				}
			
			conexion = new OleDbConnection(cad_con);//creamos la conexion con la hoja de excel o Access
			conexion.Open(); //abrimos la conexion
		}
		
		public void consultar(){
			dataAdapter = new OleDbDataAdapter(sql, conexion); //traemos los datos de la hoja y las guardamos en un dataSdapter
			dataSet = new DataSet(); // creamos la instancia del objeto DataSet
			dataAdapter.Fill(dataSet, nom_tabla);//llenamos el dataset
			dataGridView1.DataSource = dataSet.Tables[0] ; //le asignamos al DataGridView el contenido del dataSet
            tabla_excel = dataGridView1.DataSource as DataTable;
			tot_row=dataGridView1.Rows.Count;
			tot_col=dataGridView1.Columns.Count;
            dataSet.Dispose();
			//tipo = dataGridView1.Columns[0].ValueType.ToString();
			
		}
		
		public void guardar_excel(){
			
			try{
				/*Microsoft.Office.Interop.Excel.Application aplicacion;
				Microsoft.Office.Interop.Excel.Workbook libros_trabajo;
				Microsoft.Office.Interop.Excel.Worksheet hoja_trabajo;
				aplicacion = new Microsoft.Office.Interop.Excel.Application();
				libros_trabajo = aplicacion.Workbooks.Add();
				hoja_trabajo = (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajo.Worksheets.get_Item(1);
				Microsoft.Office.Interop.Excel.Range rango,rango1,rango2,rango3;
				rango = hoja_trabajo.get_Range("a2","a"+(tot_row+1));
				rango.NumberFormat = "@";
				rango3 = hoja_trabajo.get_Range("b2","b"+(tot_row+1));
				rango3.NumberFormat = "@";
				rango1 = hoja_trabajo.get_Range("d2","d"+(tot_row+1));
				rango1.NumberFormat="$ ###,###,###.##";
				rango2 = hoja_trabajo.get_Range("e2","e"+(tot_row+1));
				rango2.NumberFormat = "@";*/
				
				//MessageBox.Show("FILAS: "+tot_row+"\nCOLUMNAS"+tot_col);
				//m=1;
				/*j=0;
				
				if(n==1){
					for(j=0;j<6;j++){
						
						tabla_excel.Columns.Add(columnas[j]);
					}
				}else{
					for(j=0;j<tot_col;j++){
						Invoke(new MethodInvoker(delegate
						                         {
                                                     tabla_excel.Columns.Add(dataGridView1.Columns[j].Name.ToString());
						                         	//hoja_trabajo.Cells[i+1,j+1] = dataGridView1.Columns[j].Name.ToString();
						                         }));
					}
				}
				j=0;
				m=1;
				fila = new string[tot_col];
				
				for(i=0;i<tot_row;i++){
					m=m+1;
					
					if(n==1){
						Invoke(new MethodInvoker(delegate
						                         {
						                         	//Leer Campos
						                         	reg_pat1=Convert.ToString(dataGridView1.Rows[i].Cells[0].FormattedValue);
						                         	per=Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
						                         	t_d=Convert.ToString(dataGridView1.Rows[i].Cells[2].Value);
						                         	imp_tot=Convert.ToString(dataGridView1.Rows[i].Cells[3].Value);
						                         	fecha_pago=Convert.ToString(dataGridView1.Rows[i].Cells[4].FormattedValue);
						                         	folio=Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
						                         }));
						//Formatear el Registro Patronal
						reg_pat=(reg_pat1.Substring(1,8))+"10";
						per=(per.Substring(4,2))+"/"+(per.Substring(0,4));
						if(camb_td==1){
							t_d="6";
						}
						
						/*hoja_trabajo.Cells[m,1]=reg_pat;
						hoja_trabajo.Cells[m,2]=per;
						hoja_trabajo.Cells[m,3]=t_d;
						hoja_trabajo.Cells[m,4]=imp_tot;
						hoja_trabajo.Cells[m,5]=fecha_pago;
						hoja_trabajo.Cells[m,6]=folio;
                        tabla_excel.Rows.Add(reg_pat,per,t_d,imp_tot,fecha_pago,folio);

					}else{
   
						for(j=0;j<tot_col;j++){
							Invoke(new MethodInvoker(delegate
							                         {
							                         	valor_celda = dataGridView1.Rows[i].Cells[j].Value.ToString();
							                         	
                        	                         }));
							
							if(valor_celda!=null){
								Invoke(new MethodInvoker(delegate {
								                         	//hoja_trabajo.Cells[m,j+1]=dataGridView1.Rows[i].Cells[j].Value.ToString();
								                         	fila[j] = valor_celda;
                        		                         }));
								
							}else{
                        		fila[j]=" ";
								//hoja_trabajo.Cells[m,j+1]=" ";
							}
						}
						
						tabla_excel.Rows.Add(fila);
						
					}
					Invoke(new MethodInvoker(delegate {
					                         	progreso();
					                         }));
				}
				/*libros_trabajo.SaveAs(fichero.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
				//libros_trabajo.SaveAs(fichero.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
				libros_trabajo.Close(true);
				aplicacion.Quit();*/
				
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(tabla_excel, "hoja_lz");
                wb.SaveAs(fichero.FileName);
				Invoke(new MethodInvoker(delegate {
                                         	//label13.Text="Registros Exportados: "+(i+1)+"/"+tot_row;
											label13.Refresh();
				                         	panel1.Visible=false;
				                         	button1.Enabled=true;
				                         	button2.Enabled=true;
				                         	button3.Enabled=true;
				                         	comboBox1.Enabled=true;
				                         	checkBox1.Enabled=true;
				                         	label7.Enabled =false;
				                         	label8.Enabled =false;
				                         	label9.Enabled=true;
				                         	dataGridView1.Visible=true;
				                         	this.Text="Access to Excel";
				                         	UseWaitCursor = false;
				                         	dataGridView1.DataSource="";
				                         	label4.Text="Archivo:  ";
				                         	label10.Text="Registros: ";
				                         	checkBox1.Checked=false;
				                         	comboBox1.Items.Clear();
				                         	button4.Enabled=false;
				                         	buscarToolStripMenuItem.Enabled=false;
				                         	this.label11.Image = global::Accces_to_Excel.Properties.Resources.disconnect;
				                         }));
				MessageBox.Show("El Archivo:\n"+fichero.FileName.ToString()+"\nSe ha creado correctamente","Access to Excel - ¡Exito!");
				n=0;
				
			}catch (Exception ex)
			{
				//en caso de haber una excepcion que nos mande un mensaje de error
				MessageBox.Show("Error, Verificar el archivo o el nombre de la tabla\n"+ex,"Error al Crear Archivo de Excel");
			}
		}
		
		public void guardar_csv(){
			try{
				//open file
				StreamWriter wr = new StreamWriter(nom_save);
				
				if(n==1){
					for(j=0;j<6;j++){
						wr.Write(columnas[j] + "~");
					}
				}else{
					for(j=0;j<tot_col;j++){
						Invoke(new MethodInvoker(delegate
						                         {
						                         	wr.Write(dataGridView1.Columns[j].Name.ToString().ToUpper() + "~");
						                         }));
					}
				}

				wr.WriteLine();
				
				m=1;
				j=0;
				//write rows to excel file
				for (i = 0; i < tot_row; i++)
				{
					
					if(n==1){
						Invoke(new MethodInvoker(delegate{
						                         	//Leer Campos
						                         	
						                         	reg_pat1=Convert.ToString(dataGridView1.Rows[i].Cells[0].FormattedValue);
						                         	per=Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
						                         	t_d=Convert.ToString(dataGridView1.Rows[i].Cells[2].Value);
						                         	imp_tot=Convert.ToString(dataGridView1.Rows[i].Cells[3].Value);
						                         	fecha_pago=Convert.ToString(dataGridView1.Rows[i].Cells[4].FormattedValue);
						                         	folio=Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
						                         }));
						//Formatear el Registro Patronal
						reg_pat=(reg_pat1.Substring(1,8))+"10";
						per=(per.Substring(4,2))+"/"+(per.Substring(0,4));
						if(camb_td==1){
							t_d="6";
						}
						
						wr.Write(reg_pat + "~");
						wr.Write(per + "~");
						wr.Write(t_d + "~");
						wr.Write(imp_tot + "~");
						wr.Write(fecha_pago + "~");
						wr.Write(folio + "~");
						
					}else{
						
						for(j=0;j<tot_col;j++){
							Invoke(new MethodInvoker(delegate
							                         {
							                         	valor_celda = dataGridView1.Rows[i].Cells[j].Value.ToString();
							                         }));
							
							if(valor_celda!=null){
								Invoke(new MethodInvoker(delegate
								                         {
								                         	wr.Write(dataGridView1.Rows[i].Cells[j].Value + "~");
								                         }));
								
							}else{
								wr.Write("~");
							}
						}
						
					}
					
					wr.WriteLine();
					Invoke(new MethodInvoker(delegate
					                         {
					                         	progreso();
					                         }));

				}
				
				Invoke(new MethodInvoker(delegate {
				                         	panel1.Visible=false;
				                         	button1.Enabled=true;
				                         	button2.Enabled=true;
				                         	button3.Enabled=true;
				                         	comboBox1.Enabled=true;
				                         	checkBox1.Enabled=true;
				                         	label7.Enabled =false;
				                         	label8.Enabled =false;
				                         	label9.Enabled=true;
				                         	dataGridView1.Visible=true;
				                         	UseWaitCursor = false;
				                         	this.Text="Access to Excel";
				                         	dataGridView1.DataSource="";
				                         	label4.Text="Archivo:  ";
				                         	label10.Text="Registros: ";
				                         	checkBox1.Checked=false;
				                         	comboBox1.Items.Clear();
				                         	button4.Enabled=false;
				                         	buscarToolStripMenuItem.Enabled=false;
				                         	this.label11.Image = global::Accces_to_Excel.Properties.Resources.disconnect;
				                         }));
				
				MessageBox.Show("El Archivo:\n"+fichero.FileName.ToString()+"\nSe ha creado correctamente","¡Exito!");
				n=0;
				//close file
				wr.Close();
			}catch (Exception ex)
			{
				//en caso de haber una excepcion que nos mande un mensaje de error
				MessageBox.Show("Error, Verificar el archivo o el nombre de la tabla\n"+ex,"Error al Guardar Archivo CSV");
			}
		}
		
		public void cargar_tabla(){
			
			try{
				if((comboBox1.SelectedItem)== null){
					MessageBox.Show("Tienes que Elegir una Tabla");
					button4.Enabled=false;
					button7.Enabled=false;
					buscarToolStripMenuItem.Enabled=false;
				}else{
					nom_tabla = comboBox1.SelectedItem.ToString();
					button4.Enabled=true;
					button7.Enabled=true;
					buscarToolStripMenuItem.Enabled=true;
					if(k==1){
						
						if(((nom_file.Substring(0,4)).Equals("RCCS"))||((nom_file.Substring(0,4)).Equals("CDCS"))){
							
							if((nom_file.Substring(0,4)).Equals("RCCS")){
								if((nom_tabla.Substring(0,4)).Equals("CDRC")){
									sql= "SELECT RRC_PATRON, RRC_PER, RRC_DOC, RRC_IMP_TOT, RRC_FEC_MOV, RRC_NUM_FOL FROM "+nom_tabla+" WHERE RRC_SUB_EMI_CR = 38 AND RRC_MOD = 10 AND (RRC_DOC IN (1,2,3,6)) ";
									button2.Enabled=true;
									n=1;
									camb_td=1;
									consultar();
								}else{
									MessageBox.Show("No se Puede Aplicar el Filtro a la Tabla Seleccionada");
									checkBox1.Checked=false;
								}
							}
							
							if((nom_file.Substring(0,4)).Equals("CDCS")){
								if((nom_tabla.Substring(0,4)).Equals("CDRC")){
									sql= "SELECT RC_PATRON, RC_PER, RC_DOC, RC_IMP_TOT, RC_FEC_MOV, RC_NUM_FOL FROM "+nom_tabla+" WHERE RC_SUB_EMI_CR = 38 AND RC_MOD = 10 AND (RC_DOC IN (1,2,3,6)) ";
									button2.Enabled=true;
									n=1;
									camb_td=0;
									consultar();
								}else{
									MessageBox.Show("No se Puede Aplicar el Filtro a la Tabla Seleccionada");
									checkBox1.Checked=false;
								}
							}
						}
						else{
							MessageBox.Show("No se Puede Aplicar el Filtro a la Tabla Seleccionada");
							checkBox1.Checked=false;
						}
					}else{
						sql= "SELECT * FROM "+nom_tabla+" ";
						button2.Enabled=true;
						consultar();
					}
					tot_row=dataGridView1.Rows.Count;
					label10.Text = "Registros: "+(tot_row);
					label8.Enabled =true;
				}
			}catch (Exception ex)
			{
				//en caso de haber una excepcion que nos mande un mensaje de error
				MessageBox.Show("Error, Verificar el archivo o el nombre de la tabla\n"+ex,"Error al Abrir Archivo de Access");
			}
		}
		
		public void progreso(){
			
			por = ((i*100)/tot_row);
			if(por>99.3){
				porcent=100;
			}
			porcent = Convert.ToInt32(por);
			progressBar1.Value=porcent;
			porcent_text = Convert.ToString(porcent);
			label6.Text=porcent_text + " %";
			this.Text="Access to Excel ("+porcent_text + "%)";
			label6.Refresh();
			label13.Text="Registros Exportados: "+i+"/"+tot_row;
			label13.Refresh();
			l++;
			
			if(l==10){
				label5.Text="Procesando.";
				label5.Refresh();
			}
			
			if(l==20){
				label5.Text="Procesando..";
				label5.Refresh();
			}
			
			if(l==30){
				label5.Text="Procesando...";
				label5.Refresh();
			}
			
			if(l==40){
				label5.Text="Procesando";
				label5.Refresh();
				l=0;
			}
		}
		
		public void abrir_buscar(){
			
			string[] columnas_buscar = new string[tot_col];
			for(j=0; j < tot_col; j++){
				columnas_buscar[j] = dataGridView1.Columns[j].Name.ToString( );
			}
			buscar = new Form1();
			buscar.receptor(columnas_buscar);
			buscar.Show(this);
		}
		
		public void puente_busqueda(string[] busqueda, int rango){
			this.busqueda[0] = busqueda[0];
			this.busqueda[1] = busqueda[1];
			this.alcance_bus = rango;
			valor_a_buscar();
		}
		
		public void puente_res(int pos_actual){
			actual = pos_actual;
			marcar_res();
		}
		
		public void valor_a_buscar(){
			
			//MessageBox.Show("Valor a buscar: "+busqueda[0]+"\nAlcance de busqueda: "+busqueda[1]+" rango:"+alcance_bus);
			//MessageBox.Show("Tipo: "+tot_row+", "+tot_col);
			cord_res = new int[tot_col,tot_row];
			res_bus=0;
			
			if(alcance_bus == 0){
				t_bus=1;
			}else{
				t_bus=2;
			}
			
			if(t_bus==1){
				for(x=0;x<tot_row;x++){
					for(y=0;y<tot_col;y++){
						if((dataGridView1.Rows[x].Cells[y].FormattedValue.ToString()).Equals(busqueda[0])){
							cord_res[0,res_bus]=x;
							cord_res[1,res_bus]=y;
							res_bus++;
						}
						
					}
				}
			}
			
			if(t_bus==2){
				for(x=0;x<tot_row;x++){
					if((dataGridView1.Rows[x].Cells[(alcance_bus-1)].FormattedValue.ToString()).Equals(busqueda[0])){
						cord_res[0,res_bus]=x;
						cord_res[1,res_bus]=(alcance_bus-1);
						res_bus++;						
					}
				}
			}
			
			dataGridView1.Rows[cord_res[0,0]].Cells[cord_res[1,0]].Selected=true;
			buscar.resultados(res_bus);
		}
		
		public void marcar_res(){
			
			dataGridView1.Rows[cord_res[0,actual]].Cells[cord_res[1,actual]].Selected=true;
			
		}
		
		void Button2Click(object sender, EventArgs e)
		{
			
			fichero.Title = "Guardar archivo de Excel";
			fichero.Filter = "Archivo Excel (*.XLSX)|*.xlsx|Archivo Compatible con Excel (*.CSV)|*.csv";
			
			if(fichero.ShowDialog() == DialogResult.OK){
				
				panel1.Visible=true;
				panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(115)))), ((int)(((byte)(70)))));
				button1.Enabled=false;
				button2.Enabled=false;
				button3.Enabled=false;
				comboBox1.Enabled=false;
				checkBox1.Enabled=false;
				label9.Enabled=false;
				dataGridView1.Visible=false;
				UseWaitCursor = true;
				button7.Enabled = false;
				
				nom_save = fichero.FileName;
				if((nom_save.Substring((nom_save.Length-4),4)).Equals("xlsx")){
					hilosecundario = new Thread(new ThreadStart(guardar_excel));
					hilosecundario.Start();
					//guardar_excel();
				}
				
				if((nom_save.Substring((nom_save.Length-3),3)).Equals("csv")){
					hilosecundario = new Thread(new ThreadStart(guardar_csv));
					hilosecundario.Start();
					//guardar_csv();
				}
				
				button7.Enabled=true;
			}
		}
		
		void Button1Click(object sender, EventArgs e)
		{
			try{
				buscar.Close();
			}catch(NullReferenceException ex){
				
			}
			
			OpenFileDialog archivo = new OpenFileDialog();
			archivo.Title = "Seleccione el archivo de Access";
			archivo.Filter = "Archivo Access (*.mdb, *.accdb)|*.mdb;*.accdb";
			
			if(archivo.ShowDialog() == DialogResult.OK){
				
				nom_bd = archivo.FileName;
				nom_file = archivo.SafeFileName;
				label4.Text = "Archivo: "+nom_file;
				comboBox1.Items.Clear();
				dataGridView2.DataSource="";
				filas=0;
				fil_tabla="";
				i=0;
				ext=nom_file.Substring(nom_file.Length-3,3);
				
				
				try{
					conectar();
					System.Data.DataTable dt = conexion.GetSchema("TABLES");
					dataGridView2.DataSource = dt;
					if(dataGridView2.RowCount > 1){
						filas=(dataGridView2.RowCount)-1;
					}else{
						MessageBox.Show("Error al Capturar Tablas del Archivo\nReincie la Aplicación","Error al Abrir Archivo de Access");
					}
					do{
						fil_tabla = dataGridView2.Rows[i].Cells[3].Value.ToString();
						tam_tab = fil_tabla.Length;
						//MessageBox.Show("Tamaño Texto Tabla: "+tam_tab);
						if (!(fil_tabla).Equals("")){
							if (fil_tabla == "TABLE"){
								tabla=dataGridView2.Rows[i].Cells[2].Value.ToString();
								comboBox1.Items.Add(tabla);
							}
						}
						i++;
					}while(i<=filas);
					i=0;
					
				}catch (Exception ex)
				{
					//en caso de haber una excepcion que nos mande un mensaje de error
					MessageBox.Show("Error, Verificar el archivo o el nombre de la tabla\n"+ex,"Error al Abrir Archivo de Access");
				}
				this.label11.Image = global::Accces_to_Excel.Properties.Resources.connect;
				//MessageBox.Show("Se Ha cargado el Archivo Access Correctamente","¡Exito!");
				label7.Enabled =true;
			}
		}
		
		void ComboBox1SelectedIndexChanged(object sender, EventArgs e)
		{
			
		}
		
		void Button3Click(object sender, EventArgs e)
		{
			cargar_tabla();
		}
		
		void CheckBox1CheckedChanged(object sender, EventArgs e)
		{
			if(checkBox1.Checked){
				if((comboBox1.SelectedItem)== null){
					checkBox1.Checked=false;
					MessageBox.Show("Tienes que Elegir una Tabla");
					
				}else{
					k=1;
				}
				
			}else{
				k=0;
			}
			
		}
		
		void CheckBox1CheckStateChanged(object sender, EventArgs e)
		{
			
		}
		
		void DataGridView1CellContentClick(object sender, DataGridViewCellEventArgs e)
		{
			
		}
		
		void BuscarToolStripMenuItemClick(object sender, EventArgs e)
		{
			abrir_buscar();
		}
		
		void Button4Click(object sender, EventArgs e)
		{
			abrir_buscar();
		}
		
		void Button5Click(object sender, EventArgs e)
		{
			Form2 about = new Form2();
			about.Show();
		}
		
		void MainFormLoad(object sender, EventArgs e)
		{
			this.Top = (Screen.PrimaryScreen.WorkingArea.Height - this.Height) / 2;
			this.Left = (Screen.PrimaryScreen.WorkingArea.Width - this.Width) / 2;
			toolTip1.SetToolTip(this.button1, "Selecciona el Archivo MDB a cargar");
			toolTip1.SetToolTip(this.button2, "Seleciona la ubicacion y\nel nombre del archivo a exportar");
			toolTip1.SetToolTip(this.button3, "Cargar Tabla del Archivo MDB");
			toolTip1.SetToolTip(this.button4, "Buscar en la tabla...");
			toolTip1.SetToolTip(this.button5, "Acerca de...");
			toolTip1.SetToolTip(this.checkBox1, "Marca para aplicar el filtro de\nla subdelegacion Hidalgo");
		}
		
		void Button6Click(object sender, EventArgs e)
		{
			/*
			do{
				nom_per0=dataGridView1.Rows[j].Cells[4].Value.ToString();
				
				if((nom_per0.Substring(0,6)).Equals("OTROSD")){
					if((nom_per0.Substring(6,3)).Equals("RCV")){
						nom_per1="RCV_CM_"+(nom_per0.Substring(nom_per0.Length-3,3));
					}else{
						nom_per1="COP_CM_"+(nom_per0.Substring(nom_per0.Length-3,3));
					}
				}else{
					
					if((nom_per0.Substring(nom_per0.Length-3,3)).Equals("RCV")){
						if((nom_per0.Substring(0,6)).Equals("SIVEPA")){
							nom_per1="RCV_"+(nom_per0.Substring(0,6))+"_"+(nom_per0.Substring(6,3))+"_"+(nom_per0.Substring(11,4))+(nom_per0.Substring(9,2));
						}else{
							nom_per1="RCV_"+(nom_per0.Substring(0,3))+"_"+(nom_per0.Substring(3,3))+"_"+(nom_per0.Substring(6,6));
						}
					}else{
						if((nom_per0.Substring(0,6)).Equals("SIVEPA")){
							nom_per1="COP_"+(nom_per0.Substring(0,6))+"_"+(nom_per0.Substring(6,3))+"_"+(nom_per0.Substring(11,4))+(nom_per0.Substring(9,2));
						}else{
							nom_per1="COP_"+(nom_per0.Substring(0,3))+"_"+(nom_per0.Substring(3,3))+"_"+(nom_per0.Substring(6,6));
						}
					}
					
				}
				
				dataGridView1.Rows[j].Cells[4].Value = nom_per1;
				j++;
			}while(dataGridView1.RowCount>j);*/
		}
		
		void Button7Click(object sender, EventArgs e)
		{
			sqlform = new Form3();
			sqlform.ShowDialog(this);
			try{
				if(cad_con_sql.Length <40){
					MessageBox.Show("Proporcione datos de conexión válidos,\nde lo contrario no podrá efectuar ninguna acción.","Conexión Inválida",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
				}else{
                    conexion.Close();
					hilosecundario = new Thread(new ThreadStart(exportarsql));
					hilosecundario.Start();
					//exportarsql();
				}
			}catch{
					MessageBox.Show("Proporcione datos de conexión válidos,\nde lo contrario no podrá efectuar ninguna acción.","Conexión Inválida",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
			}
		}	
		
		public void conec(String conecs,String bd){
			cad_con_sql=conecs;
			datab_name=bd;
		}
		
		public void exportarsql(){
			String nom_per0, nom_per1="",f1,f2,f3,f4,f5,nom_tablasql,notifi="",controldr="",dias_not="",fecha_pag="",perio,nn="";
			String folio_si="",porcen="",imp_pago="",num_pago="",tdoc="",cre_cuo="",cre_mul="",reg_pat_0,reg_pat_1,reg_pat_2,fech_depu="";
			int guardar=0,no_repe=0,indes=0,jj=0;
			
			DialogResult respuesta;
			
			
			Invoke(new MethodInvoker(delegate {
			                         	
			//MessageBox.Show(cad_con_sql);

			nom_tablasql = "datos_factura_"+System.DateTime.Today.ToShortDateString().Replace('/','_');
			
			consql.conectar(cad_con_sql);
			dataGridView3.DataSource=consql.chema(datab_name);
            if (dataGridView3.RowCount > 0)
            {
                do
                {

                    if (nom_tablasql.Equals(dataGridView3.Rows[i].Cells[0].Value.ToString()))
                    {
                        if (!((nom_tablasql.Substring(nom_tablasql.Length - 3, 1)).Equals("_")))
                        {
                            nom_tablasql = nom_tablasql + "_00";
                        }
                        else
                        {
                            indes = nom_tablasql.LastIndexOf('_');
                            if (no_repe < 10)
                            {
                                nom_tablasql = nom_tablasql.Substring(0, indes + 1) + "0" + no_repe;
                            }
                            else
                            {
                                nom_tablasql = nom_tablasql.Substring(0, indes + 1) + no_repe;
                            }
                        }
                        no_repe = no_repe + 1;
                        i = 0;
                    }
                    //MessageBox.Show(nom_tablasql+","+no_repe+","+i+","+tot_tablas);
                    i++;
                } while (i < dataGridView3.RowCount);
                i = 0;
            }
            else
            {
                nom_tablasql = "datos";
            }
			
			//respuesta = MessageBox.Show(nom_tablasql);
			
			respuesta = MessageBox.Show("Se creará la tabla: \""+nom_tablasql+"\""+
			                "\n\n¿Desea Continuar?","CONFIRMAR",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
			
			if(respuesta ==DialogResult.No){
				MessageBox.Show("No se efectuó ninguna acción");
			}else{
				
				conexion.Close();
			panel1.Visible=true;
			panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(1)))), ((int)(((byte)(120)))), ((int)(((byte)(208)))));
			button1.Enabled=false;
			button2.Enabled=false;
			button3.Enabled=false;
			button4.Enabled=false;
			comboBox1.Enabled=false;
			checkBox1.Enabled=false;
			label9.Enabled=false;
			dataGridView1.Visible=false;
			UseWaitCursor = true;
			button7.Enabled = false;
			label13.Enabled = true;				
				
				
			//consql.conectar("base_principal");
			sql="CREATE TABLE "+nom_tablasql+" ("+
				"   id int(12) NOT NULL AUTO_INCREMENT,"+
				"   subdelegacion  int(2) NOT NULL DEFAULT \"0\","+
				"   incidencia  int(2) DEFAULT NULL,"+
				"   tipo_documento  int(4)  DEFAULT NULL,"+
				"   nombre_periodo  varchar(25) NOT NULL,"+
				"   registro_patronal  varchar(14) NOT NULL,"+
				"   registro_patronal1  varchar(11) NOT NULL,"+
				"   registro_patronal2  varchar(10) NOT NULL,"+
				"   razon_social  varchar(100) NOT NULL,"+
				"   periodo  varchar(6) NOT NULL,"+
				"   credito_cuotas  varchar(9) NOT NULL DEFAULT \"-\","+
				"   credito_multa  varchar(9) NOT NULL DEFAULT \"-\","+
				"   importe_cuota  decimal(10,2) NOT NULL DEFAULT \"0.00\","+
				"   importe_multa  decimal(10,2) NOT NULL DEFAULT \"0.00\","+
				"   sector_notificacion_inicial  int(2) DEFAULT NULL,"+
				"   sector_notificacion_actualizado  int(2) DEFAULT NULL,"+
				"   controlador  varchar(45) DEFAULT \"-\","+
				"   notificador  varchar(45) NOT NULL DEFAULT \"-\","+
				"   fecha_entrega  date DEFAULT NULL,"+
				"   fecha_recepcion  date DEFAULT NULL,"+
				"   fecha_notificacion  date DEFAULT NULL,"+
				"   dias_retraso  int(4) DEFAULT \"0\","+
				"   fecha_cartera  date DEFAULT NULL,"+
				"   capturado_siscob  varchar(7) NOT NULL DEFAULT \"0/0|0/0\","+
				"   status  varchar(25) NOT NULL DEFAULT \"0\","+
				"   status_credito  varchar(25) DEFAULT \"-\","+
				"   nn  varchar(15) DEFAULT \"-\","+
				"   folio_sipare_sua  varchar(512) DEFAULT \"-\","+
				"   importe_pago  decimal(10,2) DEFAULT \"0.00\","+
				"   porcentaje_pago  decimal(10,2) DEFAULT \"0.00\","+
				"   num_pago  int(2) DEFAULT \"0\","+
				"   fecha_pago  varchar(512) DEFAULT \"-\","+
				"   fecha_depuracion varchar(512) DEFAULT \"-\","+
				"   fecha_recepcion_documento  date DEFAULT NULL,"+
				"   observaciones  varchar(255) DEFAULT \"-\","+
				"   PRIMARY KEY ( id ),  UNIQUE KEY  id_UNIQUE  ( id )) ENGINE=InnoDB DEFAULT CHARSET=utf8;";
			
			consql.consultar(sql);
			
			try{
			//nom_tablasql = "datos_factura";	
			//i = 107437;
			do{
				nom_per0=dataGridView1.Rows[i].Cells[2].Value.ToString();
				
				if(nom_per0.Length > 5 ){
				
					if((nom_per0.Substring(0,6)).Equals("OTROSD")){
						if((nom_per0.Substring(6,3)).Equals("RCV")){
							nom_per1="RCV_CM_"+(nom_per0.Substring(9,nom_per0.Length-9));
						}else{
							nom_per1="COP_CM_"+(nom_per0.Substring(6,nom_per0.Length-6));
						}
					}else{
						
						if((nom_per0.Substring(nom_per0.Length-3,3)).Equals("RCV")){
							if((nom_per0.Substring(0,6)).Equals("SIVEPA")){
								nom_per1="RCV_"+(nom_per0.Substring(0,6))+"_"+(nom_per0.Substring(6,3))+"_"+(nom_per0.Substring(11,4))+(nom_per0.Substring(9,2));
							}else{
								nom_per1="RCV_"+(nom_per0.Substring(0,3))+"_"+(nom_per0.Substring(3,3))+"_"+(nom_per0.Substring(6,6));
							}
						}else{
							if((nom_per0.Substring(0,6)).Equals("SIVEPA")){
								nom_per1="COP_"+(nom_per0.Substring(0,6))+"_"+(nom_per0.Substring(6,3))+"_"+(nom_per0.Substring(11,4))+(nom_per0.Substring(9,2));
							}else{
								nom_per1="COP_"+(nom_per0.Substring(0,3))+"_"+(nom_per0.Substring(3,3))+"_"+(nom_per0.Substring(6,6));
							}
						}
						
					}
					guardar=1;
					//MessageBox.Show("1");
				}else{
				    guardar=0;
				    //MessageBox.Show("0");
				}
				
				reg_pat_0 = dataGridView1.Rows[i].Cells[4].FormattedValue.ToString();
				reg_pat_1 = dataGridView1.Rows[i].Cells[3].FormattedValue.ToString();
				reg_pat_2 = dataGridView1.Rows[i].Cells[5].FormattedValue.ToString();
				perio = dataGridView1.Rows[i].Cells[7].FormattedValue.ToString();
				cre_cuo = dataGridView1.Rows[i].Cells[8].FormattedValue.ToString();
				cre_mul = dataGridView1.Rows[i].Cells[9].FormattedValue.ToString();
				
				if((reg_pat_0.Length == 14)&&(reg_pat_1.Length == 11)&&(reg_pat_2.Length == 10)&&(perio.Length == 6)&&(cre_cuo.Length == 9)&&(cre_mul.Length == 9)){
					guardar=1;
				}else{
					guardar=0;
					
					if(((cre_mul.Length == 0)&&(cre_cuo.Length == 0))||((cre_mul.Length == 1)&&(cre_cuo.Length == 1))){
						guardar=0;
					}else{
					
						if(cre_cuo.Length < 2){
							cre_cuo="-";
							guardar=1;
						}
						
						if(cre_mul.Length < 2){
							cre_mul="-";
							guardar=1;
						}
	
						if(reg_pat_0.Length == 13){
							guardar=1;
						}
						
						if(reg_pat_1.Length == 10){
							guardar=1;
						}
					}
				}
				
				if(guardar==1){
				
				f1 = dataGridView1.Rows[i].Cells[30].FormattedValue.ToString();
				f2 = dataGridView1.Rows[i].Cells[31].FormattedValue.ToString();
				f3 = dataGridView1.Rows[i].Cells[32].FormattedValue.ToString();
				f4 = dataGridView1.Rows[i].Cells[33].FormattedValue.ToString();
				f5 = dataGridView1.Rows[i].Cells[19].FormattedValue.ToString();
				dias_not = dataGridView1.Rows[i].Cells[28].FormattedValue.ToString();
				fecha_pag = dataGridView1.Rows[i].Cells[15].FormattedValue.ToString();
				folio_si = dataGridView1.Rows[i].Cells[14].FormattedValue.ToString();
				porcen = dataGridView1.Rows[i].Cells[17].FormattedValue.ToString();
				imp_pago = dataGridView1.Rows[i].Cells[16].FormattedValue.ToString();
				num_pago = dataGridView1.Rows[i].Cells[18].FormattedValue.ToString();
				tdoc = dataGridView1.Rows[i].Cells[40].FormattedValue.ToString();
				
				//f6 = dataGridView1.Rows[i].Cells[30].FormattedValue.ToString();
				
				if(f1.Length==10){
				f1= "\""+f1.Substring(6,4)+"/"+f1.Substring(3,2)+"/"+f1.Substring(0,2)+"\"";
				}else{
					f1="NULL";
				}
				
				if(f2.Length==10){
				f2= "\""+f2.Substring(6,4)+"/"+f2.Substring(3,2)+"/"+f2.Substring(0,2)+"\"";
				}else{
					f2="NULL";
				}
				
				if(f3.Length==10){
				f3= "\""+f3.Substring(6,4)+"/"+f3.Substring(3,2)+"/"+f3.Substring(0,2)+"\"";
				}else{
					f3="NULL";
				}
				
				if(f4.Length==10){
				f4= "\""+f4.Substring(6,4)+"/"+f4.Substring(3,2)+"/"+f4.Substring(0,2)+"\"";
				}else{
					f4="NULL";
				}
				
				if(f5.Length==10){
				f5= "\""+f5.Substring(6,4)+"/"+f5.Substring(3,2)+"/"+f5.Substring(0,2)+"\"";
				}else{
					f5="NULL";
				}
				
				if(dias_not.Length < 1){
					dias_not="0";
				}
				
				if(fecha_pag.Length < 1){
					fecha_pag="-";
				}
				
				if(folio_si.Length < 1){
					folio_si="-";
				}
				
				if(porcen.Length < 1){
					porcen="0";
				}
				
				if(imp_pago.Length < 1){
					imp_pago="0";
				}
				
				if(num_pago.Length < 1){
				   num_pago="0";
				}
				
				if(tdoc.Length < 1){
					tdoc="NULL";
				}
				
				if(cre_cuo.Length < 1){
					cre_cuo="-";
				}
				
				if(cre_mul.Length < 1){
					cre_mul="-";
				}
				
				if(dataGridView1.Rows[i].Cells[39].Value != null){
					if(dataGridView1.Rows[i].Cells[39].Value.ToString().Length>0){
						nn = dataGridView1.Rows[i].Cells[39].Value.ToString();
					}else{
						nn = "-"; 
					}
				}
				
				if(dataGridView1.Rows[i].Cells[29].Value != null){
					if(dataGridView1.Rows[i].Cells[29].Value.ToString().Length>0){
						notifi = dataGridView1.Rows[i].Cells[29].Value.ToString();
					}else{
						notifi = "-"; 
					}
				}
				
				if(dataGridView1.Rows[i].Cells[38].Value != null){
					if(dataGridView1.Rows[i].Cells[38].Value.ToString().Length>0){
						controldr = dataGridView1.Rows[i].Cells[38].Value.ToString();
					}else{
						controldr = "-"; 
					}
				}
				
				if(dataGridView1.Rows[i].Cells[20].Value != null){
					if(dataGridView1.Rows[i].Cells[20].Value.ToString().Length>0){
						fech_depu = dataGridView1.Rows[i].Cells[20].Value.ToString();
					}else{
						fech_depu = "-"; 
					}
				}
				
				/*if(dataGridView1.Rows[i].Cells[29].FormattedValue.ToString().Length < 1){
					notifi="-";
				}*/
				
				sql="INSERT INTO "+nom_tablasql+" (subdelegacion,incidencia,tipo_documento,nombre_periodo,registro_patronal,registro_patronal1,registro_patronal2,razon_social,"+
				                    "periodo,credito_cuotas,credito_multa,importe_cuota,importe_multa,sector_notificacion_inicial,sector_notificacion_actualizado,"+
				                    "controlador,notificador,fecha_entrega,fecha_recepcion,fecha_notificacion,dias_retraso,fecha_cartera,capturado_siscob,status,status_credito,nn,"+
				                    "folio_sipare_sua,importe_pago,porcentaje_pago,num_pago,fecha_pago,fecha_depuracion,fecha_recepcion_documento,observaciones) VALUES "+
				                    "("+dataGridView1.Rows[i].Cells[1].FormattedValue.ToString()+",1,"+     //subdelegacion,incidencia
				                    ""+tdoc+","+                                                           //tipo_documento
				                    "\""+nom_per1+"\","+                                                   //nombre_periodo
				                    "\""+reg_pat_0+"\","+                                                  //reg_pat
				                    "\""+reg_pat_1+"\","+                                                  //reg_pat1
				                    "\""+reg_pat_2+"\","+   											   //reg_pat2
				                    "\""+dataGridView1.Rows[i].Cells[6].FormattedValue.ToString()+"\","+   //razon_social
				                    "\""+perio+"\","+                                                      //periodo
				                    "\""+cre_cuo+"\","+                                                    //credito_cuotas
				                    "\""+cre_mul+"\","+                                                    //credito_multa
				                    ""+dataGridView1.Rows[i].Cells[10].FormattedValue.ToString()+","+      //importe_cuota
				                    ""+dataGridView1.Rows[i].Cells[11].FormattedValue.ToString()+","+      //importe_multa
				                    ""+dataGridView1.Rows[i].Cells[12].FormattedValue.ToString()+","+      //sector_notificacion_inicial
				                    ""+dataGridView1.Rows[i].Cells[13].FormattedValue.ToString()+","+//sector_notificacion_actualizado,controlador				                    
				                    "\""+controldr+"\","+                                                  //controlador
									"\""+notifi+"\","+                                                     //notificador
				                    ""+f1+","+  													       //fecha_entrega
				                    ""+f2+","+ 														       //fecha_recepcion
				                    ""+f3+","+  													       //fecha_notificacion
				                    ""+dias_not+","+                                                       //dias_retraso
				                    ""+f4+",\"0/0|0/0\","+											       //fecha_cartera,capturado_siscob
				                    "\""+dataGridView1.Rows[i].Cells[34].FormattedValue.ToString()+"\","+  //status                 
									"\""+dataGridView1.Rows[i].Cells[35].FormattedValue.ToString()+"\","+  //status_credito
									"\""+nn+"\","+  													   //nn
									"\""+folio_si+"\","+                                                   //folio_sipare_sua
									""+imp_pago+","+                                                       //importe_pago
									""+porcen+","+                                                         //porcentaje_pago
									""+num_pago+","+                                                       //num_pago									
									"\""+fecha_pag+"\","+                                                  //fecha_pago
									"\""+fech_depu+"\","+                                                  //fecha_depuracion
									""+f5+",\"-\")";                                                       //fecha_recepcion_documento,observaciones
									   
                    
                     				
		       //MessageBox.Show(sql);
				consql.consultar(sql);
				sql="";
				tdoc="";
				nom_per1="";
				reg_pat_0="";
				reg_pat_1="";
				reg_pat_2="";
				perio="";                                 
				cre_cuo="";                                                 
				cre_mul="";
				controldr="";
				notifi="";
				f1="";
				f2="";
				f3="";
				dias_not="";
				f4="";
				nn="";
				folio_si="";
				imp_pago="";
				porcen="";
				num_pago="";							
				fecha_pag="";
				fech_depu  ="";
				f5  =""; 
				
				}else{
					Mensaje mem = new Mensaje("Alguno de estos campos está vacío o posee algún error:\r\nRegistro Patronal\r\nRegistro Patronal1\r\nRegistro Patronal2\r\nPeriodo\r\nCrédito Cuota\r\nCrédito Multa","Indice: "+(i+1)+"\r\nRegistro Patronal: "+reg_pat_0+"\r\nRegistro Patronal1: "+reg_pat_1+"\r\nRegistro Patronal2: "+reg_pat_2+"\r\nPeriodo: "+perio+"\r\nCrédito Cuota: "+cre_cuo+"\r\nCrédito Multa: "+cre_mul+"");
					mem.ShowDialog();
				}
				i++;
				
				
				progreso();
				
			}while(i<tot_row);
			
			panel1.Visible=false;
			button1.Enabled=true;
			button2.Enabled=true;
			button3.Enabled=true;
			button4.Enabled=true;
			comboBox1.Enabled=true;
			checkBox1.Enabled=true;
			label9.Enabled=true;
			dataGridView1.Visible=true;
			UseWaitCursor = false;
			button7.Enabled = true;
			label13.Enabled = false;
			i=0;
			consql.cerrar();
			
			}catch(Exception e){
				Mensaje men = new Mensaje(e.ToString(),("Indice:"+(i+1)+"\r\n"+sql));
				men.ShowDialog();
			}
			}
			}));
		}
		
	}
}
