using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;

namespace FileProcessor
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services

            // Web API routes
            config.MapHttpAttributeRoutes();

            //config.Routes.MapHttpRoute(
            //    name: "DefaultApi",
            //    routeTemplate: "api/{controller}/{id}",
            //    defaults: new { id = RouteParameter.Optional }
            //);

            
            config.Routes.MapHttpRoute("DefaultApiWithAction", "Api/{controller}/{action}");
            //config.Routes.MapHttpRoute("DefaultApiWithActionAndId", "Api/{controller}/{action}/{id}",
            //    defaults: new{ id = RouteParameter.Optional });

        }
    }
}
