using AttachmentService.Models;
using Net4UserTokenLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Http.Results;

namespace AttachmentService.Controllers
{
    [EnableCors("*", "*", "*")]
    public class UserController : ApiController
    {
        // POST: api/UserToken
        [HttpGet]
        public IHttpActionResult GetUser()
        {
            try
            {
                IEnumerable<string> values;
                if (Request.Headers.TryGetValues("user-token", out values) == false)
                { 
                    return new BadRequestErrorMessageResult("user-token header not present", this);
                }
                var userToken = values.FirstOrDefault();
                string userId = AttachmentService.Services.UserIdService.GetUserId(userToken);
                return Json(userId);
            }
            catch (Exception)
            {
                return new InternalServerErrorResult(this);
            }
        }
    }
}
