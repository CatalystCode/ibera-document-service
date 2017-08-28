using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Azure;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System.IO;
using AttachmentService.Models;
using System.Xml.Linq;
using System.Security.Cryptography;
using System.Net;
using System.Text;

namespace AttachmentService.Services
{
    public class AttachmentService
    {
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
        private static Stream CopyAndClose(Stream inputStream)
        {
            const int readSize = 256;
            byte[] buffer = new byte[readSize];
            MemoryStream ms = new MemoryStream();

            int count = inputStream.Read(buffer, 0, readSize);
            while (count > 0)
            {
                ms.Write(buffer, 0, count);
                count = inputStream.Read(buffer, 0, readSize);
            }
            ms.Position = 0;
            inputStream.Close();
            return ms;
        }

        // This method processes the response from the Exchange server.
        // In your application the bulk of the processing occurs here.
        private static List<AttachmentProcessingDetails> ProcessXmlResponse(XElement responseEnvelope, ServiceRequest request, string folderName, CloudBlobContainer container)
        {
            List<AttachmentProcessingDetails> al = new List<AttachmentProcessingDetails>();
            SHA256 mySHA256 = SHA256Managed.Create();
            // First, check the response for web service errors.
            var errorCodes = from errorCode in responseEnvelope.Descendants
                             ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                             select errorCode;
            // Return the first error code found.
            foreach (var errorCode in errorCodes)
            {
                if (errorCode.Value != "NoError")
                {
                    return null;
                }
            }

            // No errors found, proceed with processing the content.
            // First, get and process file attachments.
            var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                              ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                                  select fileAttachment;
            foreach (var fileAttachment in fileAttachments)
            {
                var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
                var fileName = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Name").Value;
                var fileData = System.Convert.FromBase64String(fileContent.Value);
                var s = new MemoryStream(fileData);
                // Compute the hash of the fileStream.
                byte[] hashValue;
                hashValue = mySHA256.ComputeHash(s);
                // the blobname consists of the foldername and the hash to ensure that a different file can't be overwritten 
                var blobName = folderName + "/" + HttpUtility.UrlEncode(hashValue) + "/" + fileName;
                CloudBlockBlob blockBlob = container.GetBlockBlobReference(blobName);

                s.Position = 0;
                // start from scratch again
                if (request.upload)
                {
                    blockBlob.UploadFromStream(s);
                }

                al.Add(new AttachmentProcessingDetails()
                {
                    name = fileName,
                    url = request.upload ? blockBlob.StorageUri.PrimaryUri.AbsoluteUri : "",
                    hash = Convert.ToBase64String(hashValue),
                    sasToken = request.upload ? GetBlobSasUri(container, blobName) : ""
                });
            }
            return al;

        }
        public static ServiceResponse GetAttachmentsFromExchangeServerUsingEWS(ServiceRequest request, string folderName)
        {
            var storageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("StorageAccountConnectionString"));
            var blobClient = storageAccount.CreateCloudBlobClient();
            var container = blobClient.GetContainerReference(request.containerName);
            container.CreateIfNotExists();

            var attachmentsProcessedCount = 0;
            List<AttachmentProcessingDetails> attachmentProcessingDetails = new List<AttachmentProcessingDetails>();
            foreach (var attachment in request.attachments)
            {
                // Prepare a web request object.
                HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
                webRequest.Headers.Add("Authorization",
                  string.Format("Bearer {0}", request.attachmentToken));
                webRequest.PreAuthenticate = true;
                webRequest.AllowAutoRedirect = false;
                webRequest.Method = "POST";
                webRequest.ContentType = "text/xml; charset=utf-8";

                // Construct the SOAP message for the GetAttachment operation.
                byte[] bodyBytes = Encoding.UTF8.GetBytes(
                  string.Format(GetAttachmentSoapRequest, attachment.id));
                webRequest.ContentLength = bodyBytes.Length;

                Stream requestStream = webRequest.GetRequestStream();
                requestStream.Write(bodyBytes, 0, bodyBytes.Length);
                requestStream.Close();

                // Make the request to the Exchange server and get the response.
                HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

                // If the response is okay, create an XML document from the reponse
                // and process the request.
                if (webResponse.StatusCode == HttpStatusCode.OK)
                {
                    var responseStream = webResponse.GetResponseStream();

                    var responseEnvelope = XElement.Load(responseStream);

                    // After creating a memory stream containing the contents of the 
                    // attachment, this method writes the XML document to the trace output.
                    // Your service would perform it's processing here.
                    if (responseEnvelope != null)
                    {
                        attachmentProcessingDetails.AddRange(ProcessXmlResponse(responseEnvelope, request, folderName, container));
                    }

                    // Close the response stream.
                    responseStream.Close();
                    webResponse.Close();

                }
                // If the response is not OK, return an error message for the 
                // attachment.
                else
                {

                }
            }

            var response = new ServiceResponse();
            response.attachmentsProcessed = attachmentsProcessedCount;
            response.attachmentProcessingDetails = attachmentProcessingDetails.ToArray();
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