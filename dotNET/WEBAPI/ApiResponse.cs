using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MTPAdmin.Infrastructure
{
    /// <summary>
    /// API response for the WEBAPI tier
    /// </summary>
    /// <typeparam name="T">Datatype of the response itself</typeparam>
    public class ApiResponse<T>
    {
        /// <summary>
        /// Properties
        /// </summary>
        public bool error { get; set; }
        public bool Ok { get; set; }
        public string Message { get; set; }
        public T Data { get; set; }
        public string status { get; set; }

        internal ApiResponse() { }

        /// <summary>
        /// API response
        /// </summary>
        /// <param name="data">data of the response itself</param>
        /// <param name="ok">result of the response: true or false</param>
        /// <param name="message">text message asociated</param>
        public ApiResponse(T data, bool ok = true, string message = "")
        {
            this.Data = data;
            this.Message = message;
            this.Ok = ok;
            this.error = (!ok);
            if (ok)
                this.status = "ok";
            else
                this.status = "error";
        }
    }

}
