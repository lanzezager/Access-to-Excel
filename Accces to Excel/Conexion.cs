/*
 * Creado por SharpDevelop.
 * Usuario: Lanze Zager
 * Fecha: 17/09/2015
 * Hora: 03:39 p. m.
 * 
 * Para cambiar esta plantilla use Herramientas | Opciones | Codificación | Editar Encabezados Estándar
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Data;
using MySql.Data.MySqlClient;

namespace Accces_to_Excel
{
	/// <summary>
	/// Description of Conexion.
	/// </summary>
	public class Conexion
	{
		public Conexion()
		{
		}
		
		
	   String cad_con;
	   	   
	   //Declaración de los Elementos de la BD
		MySqlDataAdapter adaptador;
		DataTable fuente,chemat;
        DataSet Datazet = new DataSet();
        MySqlConnection con;	
        MySqlConnection con1;
		
		
        public void conectar(String cad){ 

			//cad_con = @"server=localhost; userid=lanzezager; password=mario; database=base_principal";

			try
			{			
				con = new MySqlConnection(cad);
				con.Open();
				//MessageBox.Show("MySQL version : "+ con.Database);
			} catch (MySqlException ex)
			{
				MessageBox.Show("Error: "+  ex.ToString());
			}	
        }
        
        public void probar(String cad){
        	con1 = new MySqlConnection(cad);
        	con1.Open();
        	//MessageBox.Show(""+con1.State.ToString()+"");
        	con1.Close();
        }
        
        public DataTable consultar(String sql){
        	
        	adaptador = new MySqlDataAdapter(sql, con);
        	Datazet.Reset();
        	if(!(adaptador.Equals(null))){
        		fuente = new DataTable();
        		adaptador.Fill(Datazet);
        		try{
        			fuente = Datazet.Tables[0];
        		}catch(IndexOutOfRangeException ex){
        			
        		}
        	}
        	return fuente;    
       }
       
        public void cerrar()
        {
            con.Close();
        }
        
        public DataTable chema(String bd){
        	chemat = consultar("SHOW TABLES FROM "+bd);
        	return chemat;
        }
	}
}
