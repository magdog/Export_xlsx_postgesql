using ClosedXML.Excel;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.ComponentModel;

namespace Ecsport
{
    public class LogFile
    {
        StreamWriter sw;

        public LogFile(string path)
        {
            sw = new StreamWriter(path);
        }

        public void WriteLine(string str)
        {
            sw.WriteLine(str);
            sw.Flush();
        }
    }
    class Program
    {

        public static DataTable ReturnDT(string cmdText, string connStr)
        {
            //string connStr = "Server = localhost; Port = 5432; CommandTimeout = 9999990; Database = tymen; User ID = postgres; Password = 666; MaxPoolSize = 250;";
            //string connStr = "Server = 192.168.1.51; Port = 5432; CommandTimeout = 9999990; Database = chelyabinsk; User ID = postgres; Password = 1234; MaxPoolSize = 250;";
            //string connStr = "Server=172.16.10.198;Database=gkhfree;User ID = bars; Password = 123;";
            NpgsqlConnection conn = new NpgsqlConnection(connStr); // connStr это адресс сервера  Npgsql Connection устанавливает соединение с сервером  PostgreSQL.
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn); // NpgsqlCommand передаёт запрос в сервер по указанному адресу 
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd); // адаптер из многих команд: выбрать, обновить, вставить и удалить, чтобы заполнить.
            DataTable dt = new DataTable();// Представляет одну таблицу данных в памяти.
            conn.Open(); // Открывает соединение с базой данных с настройками свойств, указанными в  conn
            da.Fill(dt); // Добавляет или обновляет строки  в указанном диапазоне dt чтобы они соответствовали строкам в источнике данных da
            conn.Close(); // Закрывает соединение с базой данных с настройками свойств, указанными в  conn
            return dt; // возвращает измененую таблицу 
        }
        private static void UpdateDB(string cmdText, string connStr)
        {
            //string connStr = "Server = 192.168.1.51; Port = 5432; CommandTimeout = 9999990; Database = chelyabinsk; User ID = postgres; Password = 1234; MaxPoolSize = 250;";
            //string connStr = "Server = localhost; Port = 5432; CommandTimeout = 9999990; Database = tymen; User ID = postgres; Password = 666; MaxPoolSize = 250;";
            //string connStr = "Server=172.16.10.198;Database=gkhfree;User ID = bars; Password = 123;";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                Console.WriteLine(" успешно");
            }
            catch (Exception e)
            {
                Console.WriteLine(" не успешно " + e.Message);
            }
            finally
            {
                conn.Close();
            }
        }
        static void Main(string[] args)
        {
            try
            {
                string connStr = "Server=46.0.207.122:5433;Database=gkh_samara;User ID=bars;Password=ltkjvfytckjdjv777;";
                //  string connStr = "Server=localhost;Database=gartest;User ID=bars;Password=1234;";
                Console.WriteLine("Connecting...\n" + connStr + "\n\nChoose operation:\n1-Добавление gara в b4_fias_houses\n 2 - Добавление gara в b4_fias");
                string operation = Console.ReadLine();


                if (operation == "1") // здесь проверяется какую команжу ты ввел в консоль 
                {
                    Console.WriteLine("Операция " + operation);
                    // выводит информацию о введенной команде на консоль 
                    var wb2 = new XLWorkbook(@"C:\Users\Dmitry\source\repos\Ecsport\Ecsport\bin\Debug\gar_fias_house.xlsx"); // создаёт Excel файл на основе того файла который будет в конце пути. формат xlsx
                    Console.WriteLine("Книга открыта");
                    for (int i = 2; i <= 469387; i++) 
                    { // 

                        {
                            Console.WriteLine(i + " Операция успешна команда 1  ");
                            Guid originalGuid = Guid.NewGuid();
                            string house_guid = wb2.Worksheet(1).Row(i).Cell(1).Value.ToString();
                            Guid.TryParse(house_guid, out var house_guids);
                            string ao_guid = wb2.Worksheet(1).Row(i).Cell(2).Value.ToString();
                            Guid.TryParse(ao_guid, out var ao_guids);
                            string house_num = wb2.Worksheet(1).Row(i).Cell(3).Value.ToString();
                            string build_num = wb2.Worksheet(1).Row(i).Cell(4).Value.ToString();
                            string struc_num = wb2.Worksheet(1).Row(i).Cell(5).Value.ToString();
                            System.DateTime update_date = Convert.ToDateTime(wb2.Worksheet(1).Row(i).Cell(6).Value.ToString());
                            System.DateTime start_date = Convert.ToDateTime(wb2.Worksheet(1).Row(i).Cell(7).Value.ToString());
                            System.DateTime end_date = Convert.ToDateTime(wb2.Worksheet(1).Row(i).Cell(8).Value.ToString());
                            string upd = $@"INSERT INTO b4_fias_house(house_guid,ao_guid,house_num,build_num, struc_num,update_date,start_date,end_date)
                            VALUES('{house_guids}','{ao_guids}','{house_num}','{build_num}', '{struc_num}','{update_date}','{start_date}','{end_date}' )";
                            UpdateDB(upd, connStr);
                        }



                    }
                }

                if (operation == "2") // здесь проверяется какую команжу ты ввел в консоль 
                {
                    Console.WriteLine("Операция " + operation); // выводит информацию о введенной команде на консоль 
                    var wb2 = new XLWorkbook(@"D:\exportgar\Debug\gar_fias.xlsx"); // создаёт Excel файл на основе того файла который будет в конце пути. формат xlsx
                    Console.WriteLine("Книга открыта");
                    for (int i = 2; i <= 100; i++)
                    {
                        int level = Convert.ToInt32(wb2.Worksheet(1).Row(i).Cell(1).Value.ToString());
                        Console.WriteLine(i + " Операция успешна команда 2  ");
                        string aoguid = wb2.Worksheet(1).Row(i).Cell(5).Value.ToString();
                        string parentguid = wb2.Worksheet(1).Row(i).Cell(6).Value.ToString();
                        string formalname = wb2.Worksheet(1).Row(i).Cell(7).Value.ToString();
                        string offname = wb2.Worksheet(1).Row(i).Cell(8).Value.ToString();
                        string shortname = wb2.Worksheet(1).Row(i).Cell(9).Value.ToString();
                        string regioncode = wb2.Worksheet(1).Row(i).Cell(10).Value.ToString();
                        string areacode = wb2.Worksheet(1).Row(i).Cell(11).Value.ToString();
                        string citycode = wb2.Worksheet(1).Row(i).Cell(12).Value.ToString();
                        string placecode = wb2.Worksheet(1).Row(i).Cell(13).Value.ToString();
                        string streetcode = wb2.Worksheet(1).Row(i).Cell(14).Value.ToString();
                        System.DateTime update_date = Convert.ToDateTime(wb2.Worksheet(1).Row(i).Cell(2).Value.ToString());
                        System.DateTime start_date = Convert.ToDateTime(wb2.Worksheet(1).Row(i).Cell(3).Value.ToString());
                        System.DateTime end_date = Convert.ToDateTime(wb2.Worksheet(1).Row(i).Cell(4).Value.ToString());

                        string upd = $@"INSERT INTO b4_fias(aolevel,updatedate,startdate,enddate,aoguid,parentguid,formalname,offname,shortname,regioncode,areacode,citycode,placecode,streetcode)
                            VALUES('{level}','{update_date}','{start_date}','{end_date}','{aoguid}','{parentguid}','{formalname}','{offname}','{shortname}', '{regioncode}','{areacode}','{citycode}','{placecode}','{streetcode}')";
                        UpdateDB(upd, connStr);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.ReadLine();
            }

            Console.ReadKey();
        }

    }
}
