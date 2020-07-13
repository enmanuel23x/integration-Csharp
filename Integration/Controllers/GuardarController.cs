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
    public class GuardarController : ApiController
    {
        Prueba[] Pruebas = new Prueba[7];

        public IEnumerable<Prueba> GetAllPruebas()
        {
            string connStr = "Server=127.0.0.1;Database=integracion;Uid=root;Pwd=toor";
            MySqlConnection conn = new MySqlConnection(connStr);
            Pruebas[0] = new Prueba { Resultados = "Todo bien\n\n" };
            //}
            //catch (Exception errosql)
            //{
            // MessageBox.Show("Error en conexion a la base de datos\n\n" + errosql.Message);
            //}
            double[] request_id = new double[50];
            double[] act_request_id = new double[50];
            string[] ms_project = new string[3];

            string act_trello_name;
            string act_init_date;
            string act_mail;
            string act_end_date;
            string act_trello_user;
            double act_estimated_hours;
            double act_time_loaded;
            double act_porcent;
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
                    //MessageBox.Show(request_id[contador1].ToString());
                    ms_project[contador1] = reader.GetString(4);
                    contador1++;
                }
                reader.Close(); //importante cerrar el reader pues solo se puede tener uno abierto a la vez
                conn.Close();
            }
            catch (Exception errosql)
            {
                Pruebas[1] = new Prueba { Resultados = "Error en la consulta\n\n" + errosql.Message };
            }
            try
            {
                conn.Open();
                MySqlDataReader reader;
                MySqlCommand command;
                string commandStr = "SELECT * FROM activities;";
                command = new MySqlCommand(commandStr, conn);
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    //MessageBox.Show("aqui\n\n");
                    //act_trello_name[contador] = reader.GetString(2);
                    act_request_id[contador1] = reader.GetDouble(1);
                    //MessageBox.Show(act_request_id[contador1].ToString());
                    contador++;
                }
                reader.Close(); //importante cerrar el reader pues solo se puede tener uno abierto a la vez
                conn.Close();
            }
            catch (Exception errosql)
            {
                Pruebas[2] = new Prueba { Resultados = "Error en la consulta1\n\n" + errosql.Message };
            }

            ArrayList tasks = new ArrayList(); // se declara array de las tareas
                                               // creamos un objeto de tipo aplicacion MSProject
            int cont = 0;
            int cont1 = 0;

            string[] task_names = new string[50];
            MSProject.Application app = null;
            app = new MSProject.Application();
            Pruebas[3] = new Prueba { arr = ms_project };
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

                            foreach (MSProject.Project proj in app.Projects)
                            {
                                //Se recorre las tareas
                                foreach (MSProject.Task task in proj.Tasks)
                                {
                                    if (task.Rollup.ToString() == "False")
                                    {
                                        act_trello_name = task.Name;
                                        act_time_loaded = task.ActualWork / 60;
                                        act_estimated_hours = task.Work / 60;
                                        act_init_date = String.Format("{0:yyyy-MM-dd HH:mm:ss}", task.Start);
                                        act_end_date = String.Format("{0:yyyy-MM-dd HH:mm:ss}", task.Finish);
                                        act_trello_user = task.ResourceNames;
                                        act_mail = task.Text10;
                                        act_porcent = task.Number10;
                                        if (cont % 10 == 0 && cont != 0)
                                        {
                                        }


                                        try
                                        {

                                            //MessageBox.Show("INSERT INTO activities (act_request_id, act_trello_name, act_init_date, act_end_date, act_estimated_hours, act_time_loaded ,act_porcent, act_title, act_trello_user, act_mail) VALUES (" + request_id[cont1] + ",'" + act_trello_name + "','" + act_init_date + "', '" + act_end_date + "', " + act_estimated_hours + "," + act_time_loaded + ", " + act_porcent + ", 'false', '" + act_trello_user + "', '" + act_mail + "')");
                                            conn.Open();
                                            MySqlDataReader reader;

                                            MySqlCommand command1;

                                            string commandStr1 = "INSERT INTO activities (act_request_id, act_trello_name, act_init_date, act_end_date, act_estimated_hours, act_time_loaded ,act_porcent, act_title, act_trello_user, act_mail) VALUES (" + request_id[cont1] + ",'" + act_trello_name + "','" + act_init_date + "', '" + act_end_date + "', " + act_estimated_hours + "," + act_time_loaded + ", " + act_porcent + ", 'false', '" + act_trello_user + "', '" + act_mail + "')";

                                            command1 = new MySqlCommand(commandStr1, conn);
                                            reader = command1.ExecuteReader();
                                            reader.Close(); //importante cerrar el reader pues solo se puede tener uno abierto a la vez
                                            conn.Close();
                                            cont++;
                                        }
                                        catch (Exception errosql)
                                        {
                                            Pruebas[4] = new Prueba { Resultados = "Error en la consulta\n\n" + errosql.Message };
                                        }


                                        //continue;
                                    }
                                    else
                                    {
                                        act_trello_name = task.Name;
                                        act_time_loaded = task.Number11 / 60;
                                        act_estimated_hours = task.Work / 60;
                                        act_init_date = String.Format("{0:yyyy-MM-dd HH:mm:ss}", task.Start);
                                        act_end_date = String.Format("{0:yyyy-MM-dd HH:mm:ss}", task.Finish);
                                        act_porcent = task.Number10;

                                        if (cont % 10 == 0 && cont != 0)
                                        {
                                        }


                                        try
                                        {
                                            //MessageBox.Show("INSERT INTO activities (act_request_id, act_trello_name, act_init_date, act_end_date, act_estimated_hours, act_time_loaded ,act_porcent, act_title, act_trello_user, act_mail) VALUES (" + request_id[cont1] + ",'" + act_trello_name + "','" + act_init_date + "', '" + act_end_date + "', " + act_estimated_hours + "," + act_time_loaded + ", " + act_porcent + ", 'true', '', '')");
                                            conn.Open();
                                            MySqlDataReader reader;

                                            MySqlCommand command1;

                                            string commandStr1 = "INSERT INTO activities (act_request_id, act_trello_name, act_init_date, act_end_date, act_estimated_hours, act_time_loaded ,act_porcent, act_title, act_trello_user, act_mail) VALUES (" + request_id[cont1] + ",'" + act_trello_name + "','" + act_init_date + "', '" + act_end_date + "', " + act_estimated_hours + "," + act_time_loaded + ", " + act_porcent + ", 'true', '', '')";

                                            command1 = new MySqlCommand(commandStr1, conn);
                                            reader = command1.ExecuteReader();
                                            reader.Close(); //importante cerrar el reader pues solo se puede tener uno abierto a la vez
                                            conn.Close();
                                            cont++;
                                        }
                                        catch (Exception errosql)
                                        {
                                            Pruebas[5] = new Prueba { Resultados = "Error en la consulta\n\n" + errosql.Message };
                                        }



                                    }

                                }
                            }


                            app.FileClose(Microsoft.Office.Interop.MSProject.PjSaveType.pjSave, false); //cerramos el fichero
                        }

                        //}
                    }
                    catch (Exception err)
                    {
                        Pruebas[6] = new Prueba { Resultados = "Error en la consulta\n\n" + err.Message };
                    }
                    cont1++;
                }
            }
            return Pruebas;
        }
    }
}
