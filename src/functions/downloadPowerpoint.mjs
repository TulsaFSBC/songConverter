import fetch from 'node-fetch';
import { apiCall } from './helperFunctions.mjs';

export async function downloadPowerpoint(requestData, accessToken, context){
    context.log("Retrieving file information...")
    const jsonFileInfo = await apiCall(`https://graph.microsoft.com/v1.0/drives/${requestData.driveId}/items/${requestData.driveItemId }`, {
        method: "GET",
        headers:{
            "Authorization": `Bearer ${accessToken}`
        }
    }) 
    const downloadUrl = jsonFileInfo.data["@microsoft.graph.downloadUrl"];
    if(await downloadUrl === undefined){
        context.error("Could not retrieve file information")
    }else{
        context.log("Retrieved file information successfully.")
    }
    context.log("Downloading file...");

    const response = await fetch(downloadUrl, {
        method: "GET",
        headers: {
            "Authorization": `Bearer ${accessToken}`
        }
    });
    if (!response.ok) {
        context.error(`Failed to download file: ${response.statusText}`);
    }
    context.log("File downloaded successfully.")
    context.log("Converting response to buffer...");
    const arrayBuffer = await response.arrayBuffer();
    const fileBuffer = Buffer.from(arrayBuffer);
    context.log(`File buffer size: ${fileBuffer.length}`);

    return {fileBuffer, jsonFileInfo};
}