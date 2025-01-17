import { apiCall } from "./helperFunctions.mjs";
import * as fs from 'fs/promises';
import env from 'env-var';

export async function uploadToSharepoint(requestData, accessToken, fileInfo, presentationFilePath, context){
    const proPresenterVersion = env.get("PRO_PRESENTER_VERSION").required().asIntPositive();
    context.log("Uploading file to SharePoint...")       
    const destinationFolder = ((fileInfo.data.parentReference.path).split("root:/")[0]) + "root:/proPresenter Files"
    context.log("Destination path: " + destinationFolder)

    const folderInfo = await apiCall(`https://graph.microsoft.com/v1.0/${destinationFolder}`, {
        method:"GET",
        headers:{
            "Authorization": `Bearer ${accessToken}`,
            "Accept": "application/json"
        }
    })              
    let fileName;
    if (proPresenterVersion == 6)  {
        fileName = (fileInfo.data.name).replace("pptx", "pro6")
    } else if(proPresenterVersion == 7){
        fileName = (fileInfo.data.name).replace("pptx", "pro")
    } else {
        context.error("Invalid ProPresenter version.")
    }                                                                                        
    const response = await apiCall(`https://graph.microsoft.com/v1.0/drives/${requestData.driveId}/items/${folderInfo.data.id}:/${fileName}:/content`, {
        method: "PUT",
        headers: {
            "Content-Type": "text/plain",
            "Authorization": `Bearer ${accessToken}`
        },
        body: await fs.readFile(presentationFilePath),
        redirect: "follow"
    })
    console.log(response)
    if(response.statusCode == 201 || response.statusCode == 200){
        context.log("File uploaded successfully.")
    }else{
        context.error(response.data.error.innerError);
    }
    return response;
}