using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MTPAdmin.Infrastructure
{
    public static class ControllerHelper
    {
        /// <summary>
        /// API Response for OK result (without data information)
        /// </summary>
        /// <param name="controller">Assosiated controller class</param>
        /// <param name="message">Specific message</param>
        /// <returns>Instance of controller data type with additional standard response information</returns>
        public static IActionResult BuildOk(this Controller controller, string message = "")
        {
            return controller.Ok(new ApiResponse<dynamic>
            {
                Data = null,
                Message = message,
                Ok = true,
                error = false,
                status = "ok"
            });
        }

        /// <summary>
        ///  API Response for OK result (with data of T datatype)
        /// </summary>
        /// <typeparam name="T">Specific datatype</typeparam>
        /// <param name="controller">Assosiated controller class</param>
        /// <param name="data">Response data of T datatype</param>
        /// <param name="message">Specific message</param>
        /// <returns>Instance of controller data type with additional standard response information</returns>
        public static IActionResult BuildOk<T>(this Controller controller, T data, string message = "")
        {
            return controller.Ok(new ApiResponse<T>
            {
                Data = data,
                Message = message,
                Ok = true,
                error = false,
                status = "ok"
            });
        }

        /// <summary>
        /// API Response for NOT OK result (without data information)
        /// </summary>
        /// <param name="controller">Assosiated controller class</param>
        /// <param name="message">Specific message</param>
        /// <returns>Instance of controller data type with additional standard response information</returns>
        public static IActionResult BuildNotOk(this Controller controller, string message = "")
        {
            return controller.Ok(new ApiResponse<dynamic>
            {
                Data = null,
                Message = message,
                Ok = false,
                error = true,
                status = "error"
            });
        }

        /// <summary>
        ///  API Response for NOT OK result (with data of T datatype)
        /// </summary>
        /// <typeparam name="T">Specific datatype</typeparam>
        /// <param name="controller">Assosiated controller class</param>
        /// <param name="data">Response data of T datatype</param>
        /// <param name="message">Specific message</param>
        /// <returns>Instance of controller data type with additional standard response information</returns>
        public static IActionResult BuildNotOk<T>(this Controller controller, T data, string message = "")
        {
            return controller.Ok(new ApiResponse<T>
            {
                Data = data,
                Message = message,
                Ok = false,
                error = true,
                status = "error"
            });
        }
    }
}
