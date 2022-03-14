/*
 * Creado por SharpDevelop.
 * Usuario: Lanze Zager
 * Fecha: 08/06/2015
 * Hora: 10:48 a. m.
 * 
 * Para cambiar esta plantilla use Herramientas | Opciones | Codificación | Editar Encabezados Estándar
 */
using System;

namespace Accces_to_Excel
{
	/// <summary>
	/// Description of Puente.
	/// </summary>
	public interface Puente
	{
		
		 void puente_busqueda(string[] busqueda, int rango);
		 
		 void puente_res(int actual);
		 
		 void conec(String conecs, String DataB);
		
	}
}
