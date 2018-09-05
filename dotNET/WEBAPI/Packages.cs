using System;
using Npgsql;

namespace MTP.BR
{
    public class Packages
    {
        //Data Access
        PostgreSQL pgSQL = new PostgreSQL();
        string sqlcmd;

        /// <summary>
        /// Get the price of a package based on its internal code and payment method
        /// </summary>
        /// <param name="package">Id of the package to quote</param>
        /// <param name="yearPayment">Price per annual payment (false for monthly payment)</param>
        /// <returns>Price of the package in USD</returns>
        public double GetPrice(int package, bool? yearPayment = true)
        {
            double package_value;
            double month_value = 0;
            double year_value = 0;

            //Package information
            sqlcmd = $@"select month_value, year_value
                        from public.administration_package
                        where id = {package};";

            using (var cnn = pgSQL.GetConnection())
            using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    month_value = Convert.ToDouble(reader["month_value"]);
                    year_value = Convert.ToDouble(reader["year_value"]);
                }
            }

            //Calculate the price
            if (yearPayment == true)
                package_value = year_value * 12;
            else
                package_value = month_value;

            return package_value;
        }

        /// <summary>
        /// Calculate the amount of days to the renewal day 
        /// </summary>
        /// <param name="client">Client ID to search</param>
        /// <returns>Amount of days to the renewal or 0 if it's expired</returns>
        public double GetDaysToRenewal(int client)
        {
            double days = 0;

            //Find the current expiration date
            //(if doesn't exist, get the current date)
            sqlcmd = $@"select COALESCE(due_date, Now()::Date) as due_date 
                        from public.administration_client 
                        where id = {client}";
            DateTime due_date = pgSQL.SeachDate(sqlcmd);

            //Calculate
            days = Math.Truncate((due_date - DateTime.Now).TotalDays);
            if (days < 0)
                days = 0;

            //return amount of days to renewal
            return days;
        }

        /// <summary>
        /// Calculate the amount of the discount because of an upgrade
        /// </summary>
        /// <param name="client">Client ID to search</param>
        /// <returns>Amount of the discount</returns>
        public double GetDiscount(int client, double days)
        {
            double discount = 0;

            //Find the current package
            sqlcmd = $@"select package_id 
                        from public.administration_client 
                        where id = {client}";
            int package = pgSQL.SeachNumber(sqlcmd);

            //Searh the package cost
            sqlcmd = $@"select year_payment 
                        from public.administration_client 
                        where id = {client}";
            bool year_payment = pgSQL.SeachBool(sqlcmd);
            double amount = GetPrice(package, year_payment);

            //Calculate
            if (year_payment)
                discount = Math.Round(amount / 365 * days, 2);
            else
                discount = Math.Round(amount / 30 * days, 2);

            //return amount of the discount
            return discount;
        }

        /// <summary>
        /// Calculate the cost and expiration date of a payment
        /// </summary>
        /// <param name="client">Customer who wants to buy the package</param>
        /// <param name="package">Package to acquire</param>
        /// <param name="yearPayment">Mode of payment</param>
        /// <returns>PaymentDetails class with the detail of the calculation made</returns>
        public PaymentDetails Upgrade(int client, int package, bool yearPayment)
        {
            //Errors
            const string err000 = "(code#000) The upgrade is not possible. Please, contact us to work it out.";
            const string err001 = "(code#001) The size of the package to acquire is the same as the current one, but the internal code differs.";
            const string err002 = "(code#002) It's too soon to renew your subscription.";
            const string err003 = "(code#003) In case of package category upgrade, you should use the same payment way.";

            try
            {
                //Response
                PaymentDetails upgrade = new PaymentDetails();

                //Variables
                DateTime add_date;
                DateTime due_date = DateTime.Now.Date;
                DateTime current_date = DateTime.Now.Date;
                string licence_type = "?";
                string package_name = "?";
                int package_id = 0;
                int qty_users = 0;
                bool valid_client = false;
                bool year_payment = false;
                double month_value = 0;
                double year_value = 0;

                //Find the data of the client's current package
                sqlcmd = $@"select c.add_date, c.due_date, c.licence_type, c.package_id, c.valid_client, c.year_payment
                            ,p.month_value, p.year_value, p.name as package_name, p.qty_users
                        from public.administration_client c
                            join public.administration_package p on c.package_id = p.id
                        where c.id = {client};";

                using (var cnn = pgSQL.GetConnection())
                using (var cmd = new NpgsqlCommand(sqlcmd, cnn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        add_date = Convert.ToDateTime(reader["add_date"]).Date;
                        if (reader["due_date"] != DBNull.Value)
                            due_date = Convert.ToDateTime(reader["due_date"]).Date;
                        licence_type = Convert.ToString(reader["licence_type"]);
                        package_id = Convert.ToInt16(reader["package_id"]);
                        package_name = Convert.ToString(reader["package_name"]);
                        qty_users = Convert.ToInt16(reader["qty_users"]);
                        valid_client = Convert.ToBoolean(reader["valid_client"]);
                        year_payment = Convert.ToBoolean(reader["year_payment"]);
                        month_value = Convert.ToDouble(reader["month_value"]);
                        year_value = Convert.ToDouble(reader["year_value"]);
                    }
                }

                //Find the size of the package to acquire (to compare with the current one)
                sqlcmd = $@"select qty_users from public.administration_package where id = {package}";
                int new_qty_users = pgSQL.SeachNumber(sqlcmd);

                //Find the name of the new package to acquire
                sqlcmd = $@"select name from public.administration_package where id = {package}";
                string new_package_name = pgSQL.SeachString(sqlcmd);

                //Possible cases of upgrade
                bool found = false;

                // ----------------------------------------
                // 1. From a FREE version to a PAY one (simple case)
                // ----------------------------------------
                if (licence_type == "F")
                {
                    found = true;
                    upgrade.Detail = $@"Update of the {package_name} package from the FREE version to the PAY version starting on current date";
                    upgrade.Price = GetPrice(package, yearPayment);
                    if (yearPayment)
                    {
                        upgrade.Detail += " with annual payment way.";
                        upgrade.DueDate = current_date.AddYears(1).Date;
                    }
                    else
                    {
                        upgrade.Detail += " with monthly payment way.";
                        upgrade.DueDate = current_date.AddMonths(1).Date;
                    }
                }

                // ----------------------------------------
                // 2. From a TRIAL version to a pay version 
                // ----------------------------------------
                if (licence_type == "T")
                {
                    found = true;
                    upgrade.Detail = $@"Update of the {package_name} package from the TRIAL version to the PAY version";

                    //Take the date from the time of the trial or the current date (the largest)
                    if (current_date < due_date)
                    {
                        current_date = due_date;
                        upgrade.Detail += " starting on trial expiration date";
                    }
                    else
                    {
                        upgrade.Detail += " starting on current date";
                    }

                    upgrade.Price = GetPrice(package, yearPayment);
                    if (yearPayment)
                    {
                        upgrade.Detail += " with annual payment way.";
                        upgrade.DueDate = current_date.AddYears(1).Date;
                    }
                    else
                    {
                        upgrade.Detail += " with monthly payment way.";
                        upgrade.DueDate = current_date.AddMonths(1).Date;
                    }
                }

                // ----------------------------------------
                // 3. Upgrade from a PAY version (renewal or change of version)
                // ----------------------------------------
                if (licence_type == "P")
                {
                    if (package_id == package)
                    {
                        // ----------------------------------------
                        // 3.1. Renewal
                        // ----------------------------------------
                        //the package to acquire is the same than the current one
                        found = true;

                        upgrade.Detail = $@"Renewal of the {package_name} package";

                        //Take the date from the time of the last due or the current date (the largest)
                        if (current_date < due_date)
                        {
                            current_date = due_date;
                            upgrade.Detail += " starting on current expiration date";
                        }
                        else
                        {
                            upgrade.Detail += " starting on current date";
                        }

                        upgrade.Price = GetPrice(package, yearPayment);

                        if (yearPayment)
                        {
                            upgrade.Detail += " with annual payment way.";
                            upgrade.DueDate = current_date.AddYears(1).Date;
                        }
                        else
                        {
                            upgrade.Detail += " with monthly payment way.";
                            upgrade.DueDate = current_date.AddMonths(1).Date;
                        }

                    }
                    else
                    {
                        //the package to acquire is different from the current one (change of version): complex cases
                        if (qty_users == new_qty_users)
                        {
                            //it's an error (invalid case: can't decide if it's an upgrade or downgrade)
                            throw new Exception(err001);
                        }
                        else
                        {
                            if (qty_users < new_qty_users)
                            {
                                // ----------------------------------------
                                // 3.2. Upgrade
                                // ----------------------------------------
                                //upgrade to a bigger package
                                found = true;

                                //An upgrade must using the same payment way
                                if (year_payment != yearPayment)
                                {
                                    throw new Exception(err003);
                                }

                                upgrade.Detail = $@"Upgrade from the {package_name} package to {new_package_name} package starting now and extending the expiration date";

                                //Take the date from the time of the last due or the current date (the largest)
                                if (current_date < due_date)
                                    current_date = due_date;

                                if (yearPayment)
                                {
                                    //Upgrade with annual payment starting now
                                    upgrade.Detail += " with annual payment way";
                                    upgrade.DueDate = current_date.AddYears(1).Date;
                                }
                                else
                                {
                                    //Upgrade with monthly payment
                                    upgrade.Detail += " with monthly payment way";
                                    upgrade.DueDate = current_date.AddMonths(1).Date;
                                }

                                //Price with discount
                                double daysToDiscount = GetDaysToRenewal(client);
                                double discount = GetDiscount(client, daysToDiscount);
                                upgrade.Price = GetPrice(package, yearPayment) - discount;
                                if (daysToDiscount > 0)
                                    upgrade.Detail += " and a discount by the unused " + daysToDiscount.ToString() + " days.";
                            }
                            else
                            {
                                // ----------------------------------------
                                // 3.3. Downgrade
                                // ----------------------------------------
                                //downgrade to a smaller package
                                //change the new package just now and extend the expiration date: refunds doesn't exist
                                found = true;

                                if (GetDaysToRenewal(client) > 7)
                                    throw new Exception(err002);

                                upgrade.Detail = $@"Change of {package_name} package to {new_package_name} package starting now and extending the expiration date";

                                //Take the date from the time of the last due or the current date (the largest)
                                if (current_date < due_date)
                                    current_date = due_date;

                                upgrade.Price = GetPrice(package, yearPayment);

                                if (yearPayment)
                                {
                                    upgrade.Detail += " with annual payment way.";
                                    upgrade.DueDate = current_date.AddYears(1).Date;
                                }
                                else
                                {
                                    upgrade.Detail += " with monthly payment way.";
                                    upgrade.DueDate = current_date.AddMonths(1).Date;
                                }
                            }
                        }
                    }
                }

                if (!found)
                    throw new Exception(err000);

                //Add common information
                upgrade.YearPayment = yearPayment;
                upgrade.Package = package;
                upgrade.LicenceType = "P";
                upgrade.ConfirmationID = "n/d";

                //Return ok with the calculated upgrade
                return upgrade;
            }
            catch (Exception e)
            {
                if (e.Message.Substring(0, 6) == "(code#")
                {
                    throw e;
                }
                else
                {
                    throw new Exception(err000 + " ### " + e.Message);
                }
            }
        }
    }

    /// <summary>
    /// Model with the data of the calculated upgrade
    /// </summary>
    public class PaymentDetails
    {
        public double Price { get; set; }
        public DateTime DueDate { get; set; }
        public string Detail { get; set; }
        public int Package { get; set; }
        public Boolean YearPayment { get; set; }
        public string ConfirmationID { get; set; }
        public string LicenceType { get; set; }
    }

    public class PaymentDataInput
    {
        public int client { get; set; }
        public int package { get; set; }
        public bool yearPayment { get; set; }
        public PaymentDetails paymentDetails { get; set; }
    }
}
