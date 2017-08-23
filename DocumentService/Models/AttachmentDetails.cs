/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

namespace AttachmentService.Models
{
  public class AttachmentDetails
  {
    public string attachmentType { get; set; }
    public string contentType { get; set; }
    public string id { get; set; }
    public bool isInline { get; set; }
    public string name { get; set; }
    public int size { get; set; }
  }
}
