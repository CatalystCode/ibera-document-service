/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

using AttachmentService.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Http.Results;



namespace AttachmentService.Controllers
{
    [EnableCors("*", "*", "*")]
    public class AttachmentController : ApiController
    {
        [HttpPost]
        public IHttpActionResult PostAttachments(ServiceRequest request)
        {
            ServiceResponse response = new ServiceResponse();

            try
            {
                IEnumerable<string> values;
                if (Request.Headers.TryGetValues("user-token", out values) == false)
                {
                    return new BadRequestErrorMessageResult("user-token header not present", this);
                }
                var userToken = values.FirstOrDefault();
                string userId = AttachmentService.Services.UserIdService.GetUserId(userToken);
                string folderName = HttpUtility.UrlEncode(userId);
                response = AttachmentService.Services.AttachmentService.GetAttachmentsFromExchangeServerUsingEWS(request, folderName);
            }
            catch (Exception ex)
            {
                return new InternalServerErrorResult(this);
            }

            return Json(response);
        }
    }
}

