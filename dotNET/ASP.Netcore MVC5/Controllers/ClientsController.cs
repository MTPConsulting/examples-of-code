using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
//model
using MTPAdmin.Models.MTPAdminViewModels;
//security
using Microsoft.AspNetCore.Authorization;
//BD postgresql
using Npgsql;
//Comboboxes
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Http;

namespace MTPAdmin.Controllers
{
    [Authorize(Roles = "Administrador")]
    public class ClientsController : Controller
    {
        //Data access object
        PostgreSQL pgSQL = new PostgreSQL();
        string sqlcmd;
        //Size of each page
        int pageSize = 10;

        /// <summary>
        /// GET: Clients 
        /// </summary>
        /// <param name="sortOrder">Indicates the desired ordering of records</param>
        /// <param name="currentFilter">Indicates the filter to apply</param>
        /// <param name="searchString">Indicates the filter currently applied</param>
        /// <param name="page">Number of page to return</param>
        /// <returns></returns>
        public ActionResult Index(string sortOrder,
                                string currentFilter,
                                string searchString,
                                int? page)
        {
            Clients item;
            List<Clients> items = new List<Clients>();

            //Send parameters to the view
            //...sort
            ViewData["CompanySort"] = (String.IsNullOrEmpty(sortOrder)) ? "company_desc" : "";
            ViewData["FeAltaSort"] = (sortOrder == "fe_alta_asc") ? "fe_alta_desc" : "fe_alta_asc";
            ViewData["SchemaSort"] = (sortOrder == "schema_asc") ? "schema_desc" : "schema_asc";
            ViewData["CurrentSort"] = sortOrder;

            //...filter
            if (searchString != null)
                page = 1;   //Restore paging if you are doing a search
            else
                searchString = currentFilter;
            ViewData["CurrentFilter"] = searchString;
            if (searchString != null)
                searchString = searchString.ToUpper();

            //Set the filter for the query
            string where = " ";
            if (!String.IsNullOrEmpty(searchString))
            {
                where = $@" where upper(c.company_name) like '%{searchString}%' 
                            or upper(c.phone) like '%{searchString}%' 
                            or upper(c.db_schema) like '%{searchString}%' 
                            or upper(c.db_name) like '%{searchString}%' 
                            or upper(l.name) like '%{searchString}%' 
                            or upper(p.name) like '%{searchString}%' 
                            ";
            }

            //Set the order for the query
            string orderby = " ";
            switch (sortOrder)
            {
                case "company_desc":
                    orderby = "order by company_name desc;";
                    break;
                case "fe_alta_asc":
                    orderby = "order by add_date, company_name;";
                    break;
                case "fe_alta_desc":
                    orderby = "order by add_date desc, company_name;";
                    break;
                case "schema_asc":
                    orderby = "order by db_schema, company_name;";
                    break;
                case "schema_desc":
                    orderby = "order by db_schema desc, company_name;";
                    break;
                default:    //company asc
                    orderby = "order by company_name;";
                    break;
            }

            //Load the list
            sqlcmd = $@"select c.*, l.name as licence_type_name, p.name as package_name
                        from public.administration_client c
                            left join public.adminview_licence_types l on c.licence_type = l.id
                            left join public.administration_package p on c.package_id = p.id
                        {where} {orderby}";

            using (var cnn = pgSQL.GetConnection())
            using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    item = new Clients();
                    item.Id = Convert.ToInt32(reader["Id"]);
                    item.CompanyName = (string)reader["company_name"];
                    item.Country = (string)reader["country"];
                    item.Language = (string)reader["language"];
                    item.AddDate = Convert.ToDateTime(reader["add_date"]).Date;
                    if (reader["due_date"] != DBNull.Value)
                        item.DueDate = Convert.ToDateTime(reader["due_date"]).Date;
                    if (reader["renew_date"] != DBNull.Value)
                        item.RenewalDate = Convert.ToDateTime(reader["renew_date"]).Date;
                    item.YearPayment = Convert.ToBoolean(reader["year_payment"]);
                    item.LicenceTypeName = (string)reader["licence_type_name"];
                    if (reader["package_id"] != DBNull.Value)
                    {
                        item.PackageID = Convert.ToInt32(reader["package_id"]);
                        item.PackageName = (string)reader["package_name"];
                    }
                    item.DbSchema = (string)reader["db_schema"];
                    item.DbName = (string)reader["db_name"];
                    item.ValidClient = Convert.ToBoolean(reader["valid_client"]);

                    items.Add(item);
                    item = null;
                }
            }

            //Returns the sorted and paged selection according to the pageSize set
            return View(Paginacion<Clients>.Create(items, page ?? 1, pageSize));
        }

        // GET: Clients/Details/5
        public ActionResult Details(int id)
        {
            Clients item = new Clients();
            HistoryClients trackitem = new HistoryClients();
            List<HistoryClients> tracking = new List<HistoryClients>();

            //Returns the main register (Client)
            sqlcmd = $@"select c.*, l.name as licence_type_name, s.description as company_size_name,
                            p.name as package_name, u.username
                        from public.administration_client c
                            left join public.adminview_licence_types l on c.licence_type = l.id
                            left join public.administration_companysize s on c.company_size_id = s.id
                            left join public.administration_package p on c.package_id = p.id
                            left join public.auth_user u on c.user_id = u.id
                        where c.id = {id};";

            using (var cnn = pgSQL.GetConnection())
            using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    item.Id = Convert.ToInt32(reader["Id"]);
                    item.CompanyName = (string)reader["company_name"];
                    if (reader["phone"] != DBNull.Value)
                        item.Phone = (string)reader["phone"];
                    if (reader["address"] != DBNull.Value)
                        item.Address = (string)reader["address"];
                    item.Country = (string)reader["country"];
                    item.Language = (string)reader["language"];
                    item.AddDate = Convert.ToDateTime(reader["add_date"]).Date;
                    if (reader["due_date"] != DBNull.Value)
                        item.DueDate = Convert.ToDateTime(reader["due_date"]).Date;
                    if (reader["renew_date"] != DBNull.Value)
                        item.RenewalDate = Convert.ToDateTime(reader["renew_date"]).Date;
                    item.YearPayment = Convert.ToBoolean(reader["year_payment"]);
                    item.LicenceType = (string)reader["licence_type"];
                    item.LicenceTypeName = (string)reader["licence_type_name"];
                    if (reader["package_id"] != DBNull.Value)
                    {
                        item.PackageID = Convert.ToInt32(reader["package_id"]);
                        item.PackageName = (string)reader["package_name"];
                    }
                    item.DbSchema = (string)reader["db_schema"];
                    item.DbName = (string)reader["db_name"];
                    item.ValidClient = Convert.ToBoolean(reader["valid_client"]);
                    item.CompanySizeId = Convert.ToInt32(reader["company_size_id"]);
                    item.CompanySizeName = (string)reader["company_size_name"];
                    item.UserID = Convert.ToInt32(reader["user_id"]);
                    item.UserName = (string)reader["username"];

                    //audit information with the list of operations performed with the client
                    HistoryClients hc = new HistoryClients();
                    item.history = hc.GetList(id);

                    //payment information with the list of operations performed with the client
                    Payments py = new Payments();
                    item.payments = py.GetList(id);
                }

            }
            return View(item);
        }

        // GET: Clients/Renewal/5
        public ActionResult Renewal(int id)
        {
            {
                return RedirectToAction("Index", "PackageUpgrade", new { ClientId = id });
            }
        }

        // GET: Clients/Edit/5
        public ActionResult Edit(int id)
        {
            {
                Models.MTPAdminViewModels.Clients item = new Clients();

                //Returns the record
                sqlcmd = $@"select c.*, l.name as licence_type_name, s.description as company_size_name,
                            p.name as package_name, u.username
                        from public.administration_client c
                            left join public.adminview_licence_types l on c.licence_type = l.id
                            left join public.administration_companysize s on c.company_size_id = s.id
                            left join public.administration_package p on c.package_id = p.id
                            left join public.auth_user u on c.user_id = u.id
                        where c.id = {id};";

                using (var cnn = pgSQL.GetConnection())
                using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        item.Id = Convert.ToInt32(reader["Id"]);
                        item.CompanyName = (string)reader["company_name"];
                        if (reader["phone"] != DBNull.Value)
                            item.Phone = (string)reader["phone"];
                        if (reader["address"] != DBNull.Value)
                            item.Address = (string)reader["address"];
                        item.Country = (string)reader["country"];
                        item.Language = (string)reader["language"];
                        item.AddDate = Convert.ToDateTime(reader["add_date"]).Date;
                        if (reader["due_date"] != DBNull.Value)
                            item.DueDate = Convert.ToDateTime(reader["due_date"]).Date;
                        if (reader["renew_date"] != DBNull.Value)
                            item.RenewalDate = Convert.ToDateTime(reader["renew_date"]).Date;
                        item.YearPayment = Convert.ToBoolean(reader["year_payment"]);
                        item.LicenceType = (string)reader["licence_type"];
                        if (reader["package_id"] != DBNull.Value)
                            item.PackageID = Convert.ToInt32(reader["package_id"]);
                        item.DbSchema = (string)reader["db_schema"];
                        item.DbName = (string)reader["db_name"];
                        item.ValidClient = Convert.ToBoolean(reader["valid_client"]);
                        item.CompanySizeId = Convert.ToInt32(reader["company_size_id"]);
                        item.CompanySizeName = (string)reader["company_size_name"];
                        item.UserID = Convert.ToInt32(reader["user_id"]);
                        item.UserName = (string)reader["username"];
                    }
                }
                //Comboboxes
                LicenceTypes licenceTypes = new LicenceTypes();
                List<SelectListItem> comboTiposLicencia = licenceTypes.GetList(item.LicenceType);
                ViewBag.ComboTiposLicencia = comboTiposLicencia;

                Packages packages = new Packages();
                List<SelectListItem> comboPackages = packages.GetList(item.PackageID);
                ViewBag.ComboPackages = comboPackages;

                return View(item);
            }
        }

        // POST: Clients/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(int id, Clients item, IFormCollection form)
        {
            try
            {
                //Comboboxes
                LicenceTypes licenceTypes = new LicenceTypes();
                var licenceType = form["licenceType"];
                List<SelectListItem> comboTiposLicencia = licenceTypes.GetList(licenceType);
                ViewBag.ComboTiposLicencia = comboTiposLicencia;
                ViewBag.licenceType = licenceType;

                Packages packages = new Packages();
                int package = Convert.ToInt32(form["package"]);
                List<SelectListItem> comboPackages = packages.GetList(package);
                ViewBag.ComboPackages = comboPackages;
                ViewBag.package = package;

                //Validations
                if ((licenceType == "P" || licenceType == "T") && (!item.DueDate.HasValue))
                {
                    throw new Exception("Las licencias 'Paid' o 'Trial' requieren fecha de vencimiento");
                }
                if (licenceType == "F" && item.DueDate.HasValue)
                {
                    item.DueDate = null;
                }

                //Client Update
                sqlcmd = $@"update public.administration_client
                            set phone = '{item.Phone}'
                                ,address = '{item.Address}'
                                ,valid_client = {item.ValidClient}
                                ,add_date = {pgSQL.DateToSQL(item.AddDate)}
                                ,due_date = {pgSQL.DateToSQL(item.DueDate)}
                                ,renew_date = {pgSQL.DateToSQL(item.RenewalDate)}
                                ,licence_type = '{licenceType}'
                                ,year_payment = {item.YearPayment}
                                ,package_id = {package}
                                ,db_schema = '{item.DbSchema.ToLower()}'
                                ,db_name = '{item.DbName.ToLower()}'
                            where id = {item.Id};";

                using (var cnn = pgSQL.GetConnection())
                using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.ExecuteNonQuery();
                }

                //Audit information
                sqlcmd = $@"insert into public.administration_historyclient
                                (date, description, client_id)
                            values({pgSQL.DateToSQL(DateTime.Now.Date)}
                                ,'Edición manual de datos por el operador <{User.Identity.Name}>: valid_client <{item.ValidClient}>, licence_type <{licenceType}>, package_id <{package}>, due_date <{item.DueDate}>, renew_date <{item.RenewalDate}>'
                                ,{item.Id});";

                using (var cnn = pgSQL.GetConnection())
                using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.ExecuteNonQuery();
                }

                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                string error = $"Error: {ex.Message.ToString()}";
                ViewBag.Error = error;
                return View(item);
            }
        }

        // GET: Clients/Delete/5
        public ActionResult Delete(int id)
        {
            {
                Models.MTPAdminViewModels.Clients item = new Clients();

                //Returns the record
                sqlcmd = $@"select * 
                        from public.administration_client
                        where id = {id};";

                using (var cnn = pgSQL.GetConnection())
                using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        item.Id = Convert.ToInt32(reader["Id"]);
                        item.CompanyName = (string)reader["company_name"];
                        if (reader["phone"] != DBNull.Value)
                            item.Phone = (string)reader["phone"];
                        item.Country = (string)reader["country"];
                        item.Language = (string)reader["language"];
                        item.AddDate = Convert.ToDateTime(reader["add_date"]).Date;
                        if (reader["due_date"] != DBNull.Value)
                            item.DueDate = Convert.ToDateTime(reader["due_date"]).Date;
                        if (reader["renew_date"] != DBNull.Value)
                            item.RenewalDate = Convert.ToDateTime(reader["renew_date"]).Date;
                        item.YearPayment = Convert.ToBoolean(reader["year_payment"]);
                        item.LicenceType = (string)reader["licence_type"];
                        if (reader["package_id"] != DBNull.Value)
                            item.PackageID = Convert.ToInt32(reader["package_id"]);
                        item.DbSchema = (string)reader["db_schema"];
                        item.DbName = (string)reader["db_name"];
                        item.ValidClient = Convert.ToBoolean(reader["valid_client"]);
                        item.CompanySizeId = Convert.ToInt32(reader["company_size_id"]);
                        item.UserID = Convert.ToInt32(reader["user_id"]);

                        //Validations
                        ViewBag.Error = "";
                        //...if it is not active, you can not block it 
                        if (!item.ValidClient)
                        {
                            ViewBag.Error = "No puede bloquear a un cliente que ya está inactivo.";
                        }

                    }
                }
                return View(item);
            }
        }

        // POST: Clients/Delete/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id, Clients item)
        {
            try
            {
                //Disable the client for access to the platform
                sqlcmd = $@"update public.administration_client
                            set valid_client = false
                            where id = {item.Id};";

                using (var cnn = pgSQL.GetConnection())
                using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.ExecuteNonQuery();
                }

                //Audit information
                sqlcmd = $@"insert into public.administration_historyclient
                                (date, description, client_id)
                            values({pgSQL.DateToSQL(DateTime.Now.Date)}
                                ,'Bloqueo manual del cliente por el operador <{User.Identity.Name}>'
                                ,{item.Id});";

                using (var cnn = pgSQL.GetConnection())
                using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.ExecuteNonQuery();
                }

                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                string error = $"Error: {ex.Message.ToString()}";
                ViewBag.Error = error;
                return View(item);
            }
        }

    }
}