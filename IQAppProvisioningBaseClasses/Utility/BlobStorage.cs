using System;
using System.IO;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;

namespace IQAppProvisioningBaseClasses
{
    public class BlobStorage
    {
        private readonly string _storageAccountKey;
        private readonly string _storageAccountName;

        //this is the only public constructor; can't use this class without this info
        public BlobStorage(string storageAccountName, string storageAccountKey, string containerName)
        {
            _storageAccountName = storageAccountName;
            _storageAccountKey = storageAccountKey;

            CloudBlobContainer = SetUpContainer(storageAccountName, storageAccountKey, containerName);
        }

        //these variables are used throughout the class
        private CloudBlobContainer CloudBlobContainer { get; }

        private CloudBlobContainer SetUpContainer(string storageAccountName, string storageAccountKey,
            string containerName)
        {
            var connectionString =
                $@"DefaultEndpointsProtocol=https;AccountName={storageAccountName};AccountKey={storageAccountKey}";

            //get a reference to the container where you want to put the files
            var cloudStorageAccount = CloudStorageAccount.Parse(connectionString);
            var cloudBlobClient = cloudStorageAccount.CreateCloudBlobClient();
            var cloudBlobContainer = cloudBlobClient.GetContainerReference(containerName);
            cloudBlobContainer.CreateIfNotExists();
            return cloudBlobContainer;
        }

        public void UploadFromFile(string localFilePath, string blobUrl)
        {
            var blob = CloudBlobContainer.GetBlockBlobReference(blobUrl);
            blob.UploadFromFile(localFilePath);
        }

        public void UploadText(string textToUpload, string targetFileName)
        {
            var blob = CloudBlobContainer.GetBlockBlobReference(targetFileName);
            blob.UploadText(textToUpload);
        }
        
        public void UploadFromByteArray(byte[] uploadBytes, string targetFileName)
        {
            var blob = CloudBlobContainer.GetBlockBlobReference(targetFileName);
            blob.UploadFromByteArray(uploadBytes, 0, uploadBytes.Length);
        }

        public byte[] DownloadToByteArray(string targetFileName)
        {
            var blob = CloudBlobContainer.GetBlockBlobReference(targetFileName);
            blob.FetchAttributes();
            var blobLength = blob.Properties.Length;
            var bytes = new byte[blobLength];
            blob.DownloadToByteArray(bytes, 0);
            return bytes;
        }

        public string DownloadText(string blobName)
        {
            var blob = CloudBlobContainer.GetBlockBlobReference(blobName);
            return blob.DownloadText();
        }

        public void DeleteBlob(string targetFileName)
        {
            var blob = CloudBlobContainer.GetBlockBlobReference(targetFileName);

            if (blob.Exists()) blob.Delete();
        }

        public void CopyContainer(string destinationContainer)
        {
            var targetContainer = SetUpContainer(_storageAccountName, _storageAccountKey, destinationContainer);

            foreach (var iBlob in CloudBlobContainer.ListBlobs(useFlatBlobListing: true))
            {
                var blob = (CloudBlockBlob)iBlob;
                var targetBlob = targetContainer.GetBlockBlobReference(blob.Name);
                targetBlob.StartCopy(blob);
            }
        }

        public void DeleteContainer()
        {
            CloudBlobContainer.Delete();
        }

        public DateTime? GetBlobLastModified(string targetFileName)
        {
            var blob = CloudBlobContainer.GetBlockBlobReference(targetFileName);

            if (!blob.Exists()) return null;

            blob.FetchAttributes();

            return blob.Properties.LastModified?.DateTime;
        }
    }
}