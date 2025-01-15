import { apiCall } from "./helperFunctions.mjs";
import * as fs from 'fs';
//trigger
export async function uploadToSharepoint(requestData, accessToken, fileInfo, presentationFilePath, context){
    context.log("Uploading file to SharePoint...")       
    const destinationFolder = ((fileInfo.data.parentReference.path).split("root:/")[0]) + "root:/proPresenter Files"
    context.log("Destination path: " + destinationFolder)

    const folderInfo = await apiCall(`https://graph.microsoft.com/v1.0/${destinationFolder}`, {
        method:"GET",
        headers:{
            "Authorization": `Bearer ${accessToken}`,
            "Accept": "application/json"
        }
    })                                                                                                              /*CHANGE THIS TO ACTUAL FILE NAME!*/   
    const response = await apiCall(`https://graph.microsoft.com/v1.0/drives/${requestData.driveId}/items/${folderInfo.data.id}:/${fileInfo.data.name}:/content`, {
        method: "PUT",
        headers: {
            "Content-Type": "text/plain",
            "Authorization": `Bearer ${accessToken}`
        },
        body: fs.readFileSync(presentationFilePath),
        redirect: "follow"
    })
    console.log(response)
    if(response.statusCode == 201 || response.statusCode == 200){
        context.log("File uploaded successfully.")
    }else{
        context.error(response.data.error.innerError);
    }
    
}