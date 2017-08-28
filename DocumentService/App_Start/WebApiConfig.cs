/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;

namespace AttachmentService
{
  public static class WebApiConfig
  {
    public static void Register(HttpConfiguration config)
    {
       config.EnableCors();
       config.Routes.MapHttpRoute(
          name: "DefaultApi",
          routeTemplate: "api/{controller}/{id}",
          defaults: new { id = RouteParameter.Optional }
      );
    }
  }
}
