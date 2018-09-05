using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MTP.BR;
using MTPAdmin.Infrastructure;

namespace MTPAdmin.Controllers
{
    [Produces("application/json")]
    [Route("api/WebServices")]
    public class WebServicesController : Controller
    {
        /// <summary>
        /// Calculate the cost and expiration date of a payment
        /// </summary>
        /// <param name="data">class instance of PaymentDataInput with the information bellow (*)</param>
        /// <param name="client">(*) Customer who wants to buy the package</param>
        /// <param name="package">(*) Package to acquire</param>
        /// <param name="yearPayment">(*) Mode of payment</param>
        /// <returns>PaymentDetails class with the detail of the calculation made</returns>
        [HttpPost("[action]")]
        public IActionResult Upgrade([FromBody]PaymentDataInput data)
        {
            try
            {
                PaymentDetails paymentDetails = new PaymentDetails();
                Packages packages = new Packages();
                paymentDetails = packages.Upgrade(data.client, data.package, data.yearPayment);

                return this.BuildOk<PaymentDetails>(paymentDetails);
            }
            catch (Exception ex)
            {
                return this.BuildNotOk(ex, ex.Message);
            }
        }

        /// <summary>
        /// Save a new payment of a subscription
        /// </summary>
        /// <param name="data">class instance of PaymentDataInput with the information bellow (*)</param>
        /// <param name="client">(*) Customer who made the payment</param>
        /// <param name="paymentDetails">(*) PaymentDetails class with the detail of the payment made (as Upgrade method returned it)</param>
        /// <returns>true for success, exception for error</returns>
        [HttpPost("[action]")]
        public IActionResult SavePayment([FromBody]PaymentDataInput data)
        {
            try
            {
                Clients clients = new Clients();
                Boolean result = clients.SavePayment(data.client, data.paymentDetails);

                return this.BuildOk("payment saved");
            }
            catch (Exception ex)
            {
                return this.BuildNotOk(ex, ex.Message);
            }
        }
    }
}