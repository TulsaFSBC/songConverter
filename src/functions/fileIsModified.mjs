import { app } from "@azure/functions";
import { getAccessToken } from './getAccessToken.mjs'
import { downloadPowerpoint } from "./downloadPowerpoint.mjs";
import { extractText } from "./extractText.mjs";
import { convertToPresentation } from "./convertToPresentation.mjs";
import { uploadToSharepoint } from "./uploadToSharepoint.mjs";
import { receiveRequest, sleep } from "./helperFunctions.mjs";

app.http('fileIsModified', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
      try{
        const requestData = await receiveRequest(request, context);
        await sleep(500);
        const msAccessToken = await getAccessToken(context);
        await sleep(500);
        const powerPoint = await downloadPowerpoint(requestData, msAccessToken, context);
        await sleep(3000);
        const textData = await extractText(powerPoint.fileBuffer, context);
        await sleep(2000);
        const presentationFilePath = await convertToPresentation(textData, context);
        await sleep(3000);
        const response = await uploadToSharepoint(requestData, msAccessToken, powerPoint.fileInfo, presentationFilePath, context);
        await sleep(4000);
        return {body: await response.data};
      }
      catch(err){
        context.error(err)
      }
    },
});