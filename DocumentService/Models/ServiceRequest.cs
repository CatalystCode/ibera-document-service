/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

namespace AttachmentService.Models
{
  public class ServiceRequest
  {
    public bool upload { get; set; }
    public string containerName { get; set; }
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public AttachmentDetails[] attachments { get; set; }
  }
}
