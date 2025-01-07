import { apiCall } from "./helperFunctions.mjs";
import * as fs from 'fs';

export async function uploadToSharepoint(requestData, accessToken, context, jsonFileInfo, outputFilePath){
    context.log("Uploading file to SharePoint...")       
    const destinationFolder = ((jsonFileInfo.data.parentReference.path).split("root:/")[0]) + "root:/proPresenter Files"
    context.log("Destination path: " + destinationFolder)

    const folderInfo = await apiCall(`https://graph.microsoft.com/v1.0/${destinationFolder}`, {
        method:"GET",
        headers:{
            "Authorization": `Bearer ${accessToken}`,
            "Accept": "application/json"
        }
    })
    const response = await apiCall(`https://graph.microsoft.com/v1.0/drives/${requestData.driveId}/items/${folderInfo.data.id}:/${outputFilePath}:/content`, {
        method: "PUT",
        headers: {
            "Content-Type": "text/plain",
            "Authorization": `Bearer ${accessToken}`
        },
        body: fs.readFileSync(`./${outputFilePath}`),
        redirect: "follow"
    })
    console.log(response)
    if(response.statusCode !== 201){
        context.error(response.data.error.innerError);
    }else{
        context.log("File uploaded successfully.")
    }
    
}