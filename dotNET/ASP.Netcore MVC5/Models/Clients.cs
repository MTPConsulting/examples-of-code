using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
//BD postgresql
using Npgsql;

namespace MTPAdmin.Models.MTPAdminViewModels
{
    public class Clients
    {
        [Key]
        public int Id { get; set; }
        [Required]
        [MaxLength(100)]
        [Display(Name = "Nombre")]
        public string CompanyName { get; set; }
        [MaxLength(50)]
        [Phone()]
        [Display(Name = "Teléfono")]
        public string Phone { get; set; }
        [Required]
        [MaxLength(2)]
        [Display(Name = "País")]
        public string Country { get; set; }
        [Required]
        [MaxLength(2)]
        [Display(Name = "Idioma")]
        public string Language { get; set; }
        [MaxLength(250)]
        [Display(Name = "Domicilio")]
        public string Address { get; set; }
        [Required]
        [DataType(DataType.Date)]
        [Display(Name = "Fecha de Alta")]
        public DateTime AddDate { get; set; }
        [DataType(DataType.Date)]
        [Display(Name = "Fecha de Renovación")]
        public DateTime RenewalDate { get; set; }
        [DataType(DataType.Date)]
        [Display(Name = "Fecha de Vto.")]
        public DateTime? DueDate { get; set; }
        [MaxLength(20)]
        [Display(Name = "DB Schema")]
        public string DbSchema { get; set; }
        [MaxLength(10)]
        [Display(Name = "DB Name")]
        public string DbName { get; set; }
        [Required]
        [Display(Name = "Habilitado")]
        public Boolean ValidClient { get; set; }
        [Required]
        [MaxLength(1)]
        [Display(Name = "Tipo de Licencia")]
        public string LicenceType { get; set; }     //GetList
        [Display(Name = "Tipo de Licencia")]
        public string LicenceTypeName { get; set; }
        [Required]
        [Display(Name = "Pago Anual")]
        public Boolean YearPayment { get; set; }
        [Required]
        [Display(Name = "Tamaño Empresa")]
        public int CompanySizeId { get; set; }      //GetList
        [Display(Name = "Tamaño Empresa")]
        public string CompanySizeName { get; set; }
        [Display(Name = "Paquete")]
        public int PackageID { get; set; }      //GetList
        [Display(Name = "Paquete")]
        public string PackageName { get; set; }
        [Required]
        [Display(Name = "Usuario")]
        public int UserID { get; set; }      //GetList
        [Display(Name = "Usuario")]
        public string UserName { get; set; }
        public List<HistoryClients> history { get; set; }    //GetList
        public List<Payments> payments { get; set; }    //GetList
    }

    public class HistoryClients
    {
        //Data access object
         PostgreSQL pgSQL = new PostgreSQL();
        string sqlcmd;

        [Key]
        public int Id { get; set; }
        [Required]
        [DataType(DataType.Date)]
        [Display(Name = "Fecha")]
        public DateTime Date { get; set; }
        [Required]
        [Display(Name = "Detalle")]
        public string Description { get; set; }

        /// <summary>
        /// Returns a list of HistoryClients with all client operations
        /// </summary>
        /// <param name="client">Client to consult</param>
        /// <returns>List of HistoryClients</returns>
        public List<HistoryClients> GetList(int client)
        {
            HistoryClients item;
            List<HistoryClients> items = new List<HistoryClients>();

            //Load the list
            sqlcmd = $@"select *
                        from public.administration_historyclient
                        where client_id = {client}
                        order by date desc;";

            using (var cnn = pgSQL.GetConnection())
            using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    item = new HistoryClients();
                    item.Id = Convert.ToInt32(reader["Id"]);
                    item.Date = Convert.ToDateTime(reader["date"]).Date;
                    item.Description = (string)reader["description"];

                    items.Add(item);
                    item = null;
                }
            }
            //Returns the list of elements
            return items;
        }
    }

    public class Payments
    {
        //Data access object
        PostgreSQL pgSQL = new PostgreSQL();
        string sqlcmd;

        [Key]
        public int Id { get; set; }
        [Required]
        [DataType(DataType.Date)]
        [Display(Name = "Fecha")]
        public DateTime Date { get; set; }
        [Required]
        [Display(Name = "Importe (USD)")]
        public double Price { get; set; }
        [Display(Name = "Detalle")]
        public string Detail { get; set; }
        [DataType(DataType.Date)]
        [Display(Name = "Fecha de Vto.")]
        public DateTime? DueDate { get; set; }
        [Required]
        [Display(Name = "Cliente")]
        public int ClientId { get; set; }
        [Required]
        [Display(Name = "Paquete")]
        public int PackageId { get; set; }
        [Required]
        [Display(Name = "Pago Anual")]
        public Boolean YearPayment { get; set; }


        /// <summary>
        /// Returns a list of Payments with all client operations
        /// </summary>
        /// <param name="client">Client to consult</param>
        /// <returns>List of Payments</returns>
        public List<Payments> GetList(int client)
        {
            Payments item;
            List<Payments> items = new List<Payments>();

            //Load the list
            sqlcmd = $@"select *
                        from public.administration_payment
                        where client_id = {client}
                        order by date desc;";

            using (var cnn = pgSQL.GetConnection())
            using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    item = new Payments();
                    item.Id = Convert.ToInt32(reader["Id"]);
                    item.Date = Convert.ToDateTime(reader["date"]).Date;
                    item.Price = Convert.ToDouble(reader["price"]);
                    item.Detail = (string)reader["detail"];
                    item.DueDate = Convert.ToDateTime(reader["due_date"]).Date;
                    item.ClientId = Convert.ToInt32(reader["client_id"]);
                    item.PackageId = Convert.ToInt32(reader["package_id"]);
                    item.YearPayment = Convert.ToBoolean(reader["year_payment"]);

                    items.Add(item);
                    item = null;
                }
            }
            //Returns the list of elements
            return items;
        }
    }
}
