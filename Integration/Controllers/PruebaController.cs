using Integration.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using MSProject = Microsoft.Office.Interop.MSProject;
using MySql.Data.MySqlClient;
using System.Collections;

namespace Integration.Controllers
{
    public class PruebaController : ApiController
    {
        Prueba[] Pruebas = new Prueba[3];

        public IEnumerable<Prueba> GetAllPruebas()
        {
            /* try
            {*/
            //cadenas de construcción de conexión a servidor
            string connStr = "Server=127.0.0.1;Database=integracion;Uid=root;Pwd=toor";
            MySqlConnection conn = new MySqlConnection(connStr);
            /*}
            catch (Exception errosql)
            {
             MessageBox.Show("Error en conexion a la base de datos\n\n" + errosql.Message);
            }*/
            double[] request_id = new double[50];
            string[] ms_project = new string[50];
            string[] act_trello_name = new string[50];
            DateTime[] act_init_date = new DateTime[50];
            DateTime[] act_init_real_date = new DateTime[50];
            DateTime[] act_end_date = new DateTime[50];
            DateTime[] act_real_end_date = new DateTime[50];
            double[] act_estimated_hours = new double[50];
            double[] act_time_loaded = new double[50];
            double[] act_porcent = new double[50];
            int contador = 0;
            int contador1 = 0;
            try
            {
                conn.Open();
                MySqlDataReader reader;
                MySqlCommand command;
                string commandStr = "SELECT * FROM request WHERE req_cargar='true';";
                command = new MySqlCommand(commandStr, conn);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    //MessageBox.Show("aqui\n\n");
                    request_id[contador1] = reader.GetDouble(0);
                    ms_project[contador1] = reader.GetString(4);
                    contador1++;
                }
                reader.Close(); //importante cerrar el reader pues solo se puede tener uno abierto a la vez
                conn.Close();
            }
            catch (Exception errosql)
            {
                Pruebas[0] = new Prueba { Resultados = "Error en la consulta\n\n" + errosql.Message };
            }
            try
            {
                conn.Open();
                MySqlDataReader reader;
                MySqlCommand command;
                string commandStr = "SELECT * FROM activities WHERE act_title = 'false';";
                command = new MySqlCommand(commandStr, conn);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    //MessageBox.Show("aqui\n\n");
                    act_trello_name[contador] = reader.GetString(2);
                    //MessageBox.Show(act_trello_name[contador]);
                    act_init_date[contador] = reader.GetDateTime(5);

                    act_init_real_date[contador] = reader.GetDateTime(6);

                    act_end_date[contador] = reader.GetDateTime(7);

                    act_real_end_date[contador] = reader.GetDateTime(8);

                    act_estimated_hours[contador] = reader.GetDouble(9);

                    act_time_loaded[contador] = reader.GetDouble(17);
                    //MessageBox.Show(act_time_loaded[contador].ToString());
                    act_porcent[contador] = reader.GetDouble(19);
                    //MessageBox.Show(act_porcent[contador].ToString());
                    contador++;
                }
                reader.Close(); //importante cerrar el reader pues solo se puede tener uno abierto a la vez
                conn.Close();
            }
            catch (Exception errosql)
            {
                Pruebas[1] = new Prueba { Resultados = "Error en la consulta\n\n" + errosql.Message };
            }

            ArrayList tasks = new ArrayList(); // se declara array de las tareas
                                               // creamos un objeto de tipo aplicacion MSProject
            MSProject.Application app = null;
            app = new MSProject.Application();
            int cont = 0;
            foreach (String project in ms_project)
            {
                if (project != null)
                {

                    try
                    {
                        // Si no hay problemas para abrir el project entrará en la condición
                        // Fijense en la info que da FileOpen pues aqui indicarás especificas como lo quieres abrir (escritura/lectura) y la ruta, como está aqui es de la forma que se pueda escribir y leer en él
                        if (app.FileOpen("C:/Home/Intelix/Mayoreo/00-Control-Solicitudes/" + project + "", false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, MSProject.PjPoolOpen.pjPoolReadWrite, Type.Missing, Type.Missing, Type.Missing, Type.Missing))
                        {
                            //Se recorren los proyectos activos
                            foreach (MSProject.Project proj in app.Projects)
                            {
                                //Se recorre las tareas
                                foreach (MSProject.Task task in proj.Tasks)
                                {
                                    for (int i = 0; i < act_trello_name.Length; i++)
                                    {
                                        if (act_trello_name[i] == task.Name && task.Rollup.ToString() == "False")
                                        {
                                            task.Number11 = act_time_loaded[i];
                                            task.Start10 = act_init_real_date[i];
                                            task.Finish10 = act_real_end_date[i];
                                            task.Number10 = act_porcent[i]; //actualizamos en el project el porcentaje
                                            if (cont % 10 == 0 && cont != 0)
                                            {
                                            }
                                            cont++;
                                            continue;
                                        }

                                    }
                                }
                                app.FileClose(Microsoft.Office.Interop.MSProject.PjSaveType.pjSave, false); //cerramos el fichero
                            }
                        }
                    }
                    catch (Exception err)
                    {
                        Pruebas[2] = new Prueba { Resultados = "Error en el proyecto\n\n" + err.Message };
                    }
                }
            }
            return Pruebas;
        }
    }
}
