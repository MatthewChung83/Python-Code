# -*- coding: utf-8 -*-
"""
Created on Sat Sep 11 15:52:18 2021

@author: MATTHEW5043
"""

#%%
import os, uuid
from azure.storage.blob import BlobServiceClient, BlobClient, ContainerClient, __version__
#%%

try:
    print("Azure Blob Storage v" + __version__ + " - Python quickstart sample")

    # Quick start code goes here

except Exception as ex:
    print('Exception:')
    print(ex)
#%%
connect_str = os.getenv('AZURE_STORAGE_CONNECTION_STRING')
print(connect_str)
#%%
#%% Create the BlobServiceClient object which will be used to create a container client
blob_service_client = BlobServiceClient.from_connection_string('DefaultEndpointsProtocol=https;AccountName=ucsbarcode;AccountKey=Z1i3n9aYQ0YjIwL3RPJscUHrGjXhBUNZ6uypuZPC6CElN4oLCj4spGVo/dqoh0R5MRmIWYb/icXjWG6AdAvTOA==;EndpointSuffix=core.windows.net')

#%%Create a unique name for the container
container_name = str(uuid.uuid4())

#%%Create the container
container_client = blob_service_client.create_container(container_name)
#%%
# Create a local directory to hold blob data
local_path = "C://data"
os.mkdir(local_path)
#%%
# Create a file in the local data directory to upload and download
local_file_name = str(uuid.uuid4()) + ".txt"
upload_file_path = os.path.join(local_path, local_file_name)

# Write text to the file
file = open(upload_file_path, 'w')
file.write("Hello, World!")
file.close()
#%%
# Create a blob client using the local file name as the name for the blob
blob_client = blob_service_client.get_blob_client(container=container_name, blob=local_file_name)

print("\nUploading to Azure Storage as blob:\n\t" + local_file_name)

# Upload the created file
with open(upload_file_path, "rb") as data:
    blob_client.upload_blob(data)
#%%
#import sys
#old_stdout = sys.stdout

#log_file = open("D://TEST//message.log","w")

#sys.stdout = log_file

#print(log_file)

#sys.stdout = old_stdout

#log_file.close()
#%%