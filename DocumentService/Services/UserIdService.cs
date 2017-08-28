using Microsoft.Azure;
using Net4UserTokenLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AttachmentService.Services
{
    public class UserIdService
    {
        public static string GetUserId(string token)
        {

            var hostUri = CloudConfigurationManager.GetSetting("OutlookIntegrationHostUri");

            return Net4UserToken.GetUserId(token, hostUri);
        }
    }
}