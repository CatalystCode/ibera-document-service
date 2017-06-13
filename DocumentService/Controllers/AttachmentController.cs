/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

using AttachmentsService.Models;
using Microsoft.Azure;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Xml;

namespace AttachmentsService.Controllers
{
    [EnableCors("*", "*", "*")]
    public class AttachmentController : ApiController
    {

        public ServiceResponse PostAttachments(ServiceRequest request)
        {
            ServiceResponse response = new ServiceResponse();

            try
            {
                response = GetAttachmentsFromExchangeServer(request);
            }
            catch (Exception ex)
            {
                response.isError = true;
                response.message = ex.Message;
            }

            return response;
        }

        private static string GetBlobSasUri(CloudBlobContainer container, string blobName, string policyName = null)
        {
            string sasBlobToken;

            // Get a reference to a blob within the container.
            // Note that the blob may not exist yet, but a SAS can still be created for it.
            CloudBlockBlob blob = container.GetBlockBlobReference(blobName);

            if (policyName == null)
            {
                // Create a new access policy and define its constraints.
                // Note that the SharedAccessBlobPolicy class is used both to define the parameters of an ad-hoc SAS, and
                // to construct a shared access policy that is saved to the container's shared access policies.
                SharedAccessBlobPolicy adHocSAS = new SharedAccessBlobPolicy()
                {
                    // When the start time for the SAS is omitted, the start time is assumed to be the time when the storage service receives the request.
                    // Omitting the start time for a SAS that is effective immediately helps to avoid clock skew.
                    SharedAccessExpiryTime = DateTime.UtcNow.AddDays(7),
                    // Permissions = SharedAccessBlobPermissions.Read | SharedAccessBlobPermissions.Write | SharedAccessBlobPermissions.Create
                    Permissions = SharedAccessBlobPermissions.Read

                };

                // Generate the shared access signature on the blob, setting the constraints directly on the signature.
                sasBlobToken = blob.GetSharedAccessSignature(adHocSAS);

                Console.WriteLine("SAS for blob (ad hoc): {0}", sasBlobToken);
                Console.WriteLine();
            }
            else
            {
                // Generate the shared access signature on the blob. In this case, all of the constraints for the
                // shared access signature are specified on the container's stored access policy.
                sasBlobToken = blob.GetSharedAccessSignature(null, policyName);

                Console.WriteLine("SAS for blob (stored access policy): {0}", sasBlobToken);
                Console.WriteLine();
            }

            // Return the URI string for the container, including the SAS token.
            return blob.Uri + sasBlobToken;
        }
        // This method does the work of making an Exchange Web Services (EWS) request to get the 
        // attachments from the Exchange server. This implementation makes an individual
        // request for each attachment, and returns the count of attachments processed.
        private ServiceResponse GetAttachmentsFromExchangeServer(ServiceRequest request)
        {
            var storageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("StorageAccountConnectionString"));
            var blobClient = storageAccount.CreateCloudBlobClient();
            var container = blobClient.GetContainerReference(request.containerName);
            container.CreateIfNotExists();

            int processedCount = 0;
            List<AttachmentProcessingDetails> attachmentProcessingDetails = new List<AttachmentProcessingDetails>();
            SHA256 mySHA256 = SHA256Managed.Create();

            foreach (AttachmentDetails attachment in request.attachments)
            {
                // Prepare a web request object.
                HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
                webRequest.Headers.Add("Authorization", string.Format("Bearer {0}", request.attachmentToken));
                webRequest.PreAuthenticate = true;
                webRequest.AllowAutoRedirect = false;
                webRequest.Method = "POST";
                webRequest.ContentType = "text/xml; charset=utf-8";

                // Construct the SOAP message for the GetAttchment operation.
                byte[] bodyBytes = Encoding.UTF8.GetBytes(string.Format(GetAttachmentSoapRequest, attachment.id));
                webRequest.ContentLength = bodyBytes.Length;

                Stream requestStream = webRequest.GetRequestStream();
                requestStream.Write(bodyBytes, 0, bodyBytes.Length);
                requestStream.Close();

                // Make the request to the Exchange server and get the response.
                HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

                // If the response is okay, create an XML document from the
                // response and process the request.
                if (webResponse.StatusCode == HttpStatusCode.OK)
                {
                    var blobName = request.folderName + "/" + attachment.name;
                    CloudBlockBlob blockBlob = container.GetBlockBlobReference(blobName);
                    byte[] hashValue;
                    // Create or overwrite the "myblob" blob with contents from a local file.
                    using (var responseStream = webResponse.GetResponseStream())
                    {
                        blockBlob.UploadFromStream(responseStream);
                        // Be sure it's positioned to the beginning of the stream.
                        responseStream.Position = 0;
                        // Compute the hash of the fileStream.
                         hashValue = mySHA256.ComputeHash(responseStream);
                    }
                    webResponse.Close();
                    processedCount++;
                    attachmentProcessingDetails.Add(new AttachmentProcessingDetails()
                    {
                        name = attachment.name,
                        url = blockBlob.StorageUri.PrimaryUri.AbsoluteUri,
                        hash = Convert.ToBase64String(hashValue),
                        sasToken = GetBlobSasUri(container, blobName)
                    });
                }

            }
            ServiceResponse response = new ServiceResponse();
            response.attachmentProcessingDetails = attachmentProcessingDetails.ToArray();
            response.attachmentsProcessed = processedCount;
            return response;
        }

        private const string GetAttachmentSoapRequest =
    @"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";
    }
}

// *********************************************************
//
// Outlook-Add-in-Javascript-GetAttachments, https://github.com/OfficeDev/Outlook-Add-in-Javascript-GetAttachments
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************