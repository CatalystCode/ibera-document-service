/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

namespace AttachmentService.Models
{
  public class AttachmentProcessingDetails
  {
    public string url { get; set; }
    public string sasToken { get; set; }
    public string hash { get; set; }
    public string name { get; set; }
  }
  public class ServiceResponse
  {
    public int attachmentsProcessed { get; set; }
    public AttachmentProcessingDetails[] attachmentProcessingDetails { get; set; }
  }
}
