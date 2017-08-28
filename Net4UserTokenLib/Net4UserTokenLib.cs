using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Auth;
using Microsoft.Exchange.WebServices.Auth.Validation;

namespace Net4UserTokenLib
{


    public class Net4UserToken
    {
        public static string GetUserId(string token, string hostUri)
        {
            try
            {
                AppIdentityToken aiToken = (AppIdentityToken)AuthToken.Parse(token);
                aiToken.Validate(new Uri(hostUri));
                // the UserIdentification will look like this:
                // "https://outlook.office365.com/autodiscover/metadata/json/100037ffe-8130-41c6-0000-000000000000"
                var split = aiToken.UniqueUserIdentification.Split('/');
                // only return the user id - in the baove case 100037ffe-8130-41c6-0000-000000000000
                return split[split.Length-1];
            }
            catch (TokenValidationException ex)
            {
                throw new ApplicationException("A client identity token validation error occurred.", ex);
            }
        }
    }
}
