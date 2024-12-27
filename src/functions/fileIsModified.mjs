import { app } from "@azure/functions";
import * as textract from '@markell12/textract'
import * as fs from 'fs';
import { v4 as uuidv4} from 'uuid';
import * as child from 'child_process'
import path from 'path';
import env from 'env-var';
import fetch from 'node-fetch';

//todo
//error handling
//code cleanup
//error checking on all API calls 

function b64(text){
    return Buffer.from(text).toString('base64');
}

async function apiCall(url, requestOptions){
    const response = await fetch(url, requestOptions);
    const data = await response.json();
    return data;
}


app.http('fileIsModified', {
    methods: ['GET','POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        const config = {
            tenantId: env.get("TENANT_ID").required().asString(),
            clientId: env.get("CLIENT_ID").required().asString(),
            clientSecret: env.get("CLIENT_SECRET").required().asString(),
            proPresenterVersion: env.get("PROPRESENTER_VERSION").required().default(7).asIntPositive()
        }

        const requestBody = await request.text();
        const requestData = JSON.parse(requestBody)
        context.log("Received request: " + requestBody)
        context.log("Retrieving Access token...")

        const accessTokenData = await apiCall(`https://login.microsoftonline.com/${config.tenantId}/oauth2/token`,{
            method: "POST",
            headers: {
                "content-type": "application/x-www-form-urlencoded"
            },
            body: new URLSearchParams({
                "grant_type": "client_credentials",
                "client_id": config.clientId,
                "client_secret": config.clientSecret,
                "resource": "https://graph.microsoft.com"
            }),
            redirect: "follow"}
        )
        const accessToken = accessTokenData["access_token"]

        context.log("Retrieving file information...")
        const jsonFileInfo = await apiCall(`https://graph.microsoft.com/v1.0/drives/${requestData.driveId}/items/${requestData.driveItemId}`, {
            method: "GET",
            headers:{
                "Authorization": `Bearer ${accessToken}`
            }
        }) 
        const downloadUrl = jsonFileInfo["@microsoft.graph.downloadUrl"];
        context.log("Downloading file...");
        var textSlides, outputFilePath

        const response = await fetch(downloadUrl, {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${accessToken}`
            }
        });

        if (!response.ok) {
            throw new Error(`Failed to download file: ${response.statusText}`);
        } else{
            context.log("File downloaded successfully.")
        }
        context.log("Converting response to buffer...");
        const arrayBuffer = await response.arrayBuffer();
        const fileBuffer = Buffer.from(arrayBuffer);
        context.log(`File buffer size: ${fileBuffer.length}`);

        context.log("Extracting text from file...")
        let pptxText;
        try {
            pptxText = await new Promise((resolve, reject) => {
                process.on('uncaughtException', (err) => {
                    context.error('Uncaught Exception:', err);
                    reject(err);
                });

                textract.fromBufferWithMime("application/vnd.openxmlformats-officedocument.presentationml.presentation", 
                    Buffer.from(fileBuffer), 
                    {
                        "preserveLineBreaks":true,
                        "preserveOnlyMultipleLineBreaks":false
                    }, 
                    (error, text) => {
                        if (error) {
                            context.error('Textract error:', error);
                            context.error('Error stack:', error.stack);
                            reject(error);
                        } else {
                            resolve(text);
                        }
                    }
                );
            });
        } catch (textractError) {
            context.error('Error in textract block:', textractError);
            throw textractError;
        }

        
        if(pptxText != undefined){
            context.log("Text extracted from file successfully.")
        }else{
            context.error("Error extracting text from file.")
        }
        textSlides = pptxText.split("\n\n");
        if(config.proPresenterVersion == 6){
            const presentationTemplates = {
                presentationHeader: fs.readFileSync('./pro6Templates/presentationHeader.txt').toString(),
                presentationFooter: fs.readFileSync('./pro6Templates/presentationFooter.txt').toString(),
                slide: fs.readFileSync('./pro6Templates/presentationSlide.txt').toString()
            }
            outputFilePath = `${(jsonFileInfo.name).replace(".pptx", ".pro6")}`
            var pro6SlidesArray = []
            textSlides.forEach(slide => {
                if(slide != ""){
                    let b64PlainText = b64(slide);
                    let slideLines = slide.split("\n");
                    let rtfSlideLinesArray = [];
                    slideLines.forEach(line =>{
                        if (line != ""){
                            let rtfLine = ("{\\fs200\\outl0\\strokewidth-20\\strokec1\\f0 {\\cf3\\ltrch " + line + "}\\li0\\sa0\\sb0\\fi0\\qc\\par}")
                            rtfSlideLinesArray.push(rtfLine)
                        }
                    })
                    let rtfSlideLines = rtfSlideLinesArray.join("\\n")
                    let rtfSlide = ("{\\rtf1\\prortf1\\ansi\\ansicpg1252\\uc1\\htmautsp\\deff2{\\fonttbl{\\f0\\fcharset0 Times New Roman;}}{\\colortbl;\\red0\\green0\\blue0;\\red255\\green255\\blue255;\\red250\\green235\\blue215;}\\loch\\hich\\dbch\\pard\\slleading0\\plain\\ltrpar\\itap0{\\lang1033\\fs100\\outl0\\strokewidth-20\\strokec1\\f2\\cf1 \\cf1\\qc \n" + rtfSlideLines + "\n } \n }");
                    let b64RTFText = b64(rtfSlide)
                    let pro6SlideString = presentationTemplates.slide;
                    pro6SlideString = pro6SlideString.replace("$PLAIN_TEXT", b64PlainText); //plugs in actual content into slide template
                    pro6SlideString = pro6SlideString.replace("$RTF_TEXT", b64RTFText);
                    pro6SlideString = pro6SlideString.replace("$SLIDE_UUID", uuidv4());
                    pro6SlideString = pro6SlideString.replace("$TEXTBOX_UUID", uuidv4());
                    pro6SlidesArray.push(pro6SlideString)
                }
            })        

            const presentationString = presentationTemplates.presentationHeader + pro6SlidesArray.join() + presentationTemplates.presentationFooter;
            fs.writeFileSync(`./${outputFilePath}`, presentationString, err => {
                if (err) {
                    console.error(err);
                } else {
                    context.log("pro6 File created successfully.")
                }
            });
            
        } else if (config.proPresenterVersion == 7){

            const presentationTemplates = {
                presentation: fs.readFileSync('./pro7Templates/presentation.txt').toString(),
                slide: fs.readFileSync('./pro7Templates/slide.txt').toString(),
                slideText: fs.readFileSync('./pro7Templates/slideText.txt').toString(),
                textLine: fs.readFileSync('./pro7Templates/textLine.txt').toString(),
                slideIdentifier: fs.readFileSync('./pro7Templates/slideIdentifier.txt').toString()
            }

            var pro7SlidesArray = [],
                slideIdentifierGuids = [],
                slideIdentifiers = [];
            outputFilePath = `${(jsonFileInfo.name).replace(".pptx", ".pro")}`
            
            textSlides.forEach(slide => {
                if(slide != ""){
                    let slideId = uuidv4();
                    let slideLines = slide.split("\n");
                    let rtfSlideLinesArray = [];
                    slideLines.forEach(line =>{
                        if (line != ""){
                            let rtfLine = presentationTemplates.textLine.replace("$TEXT", line)
                            rtfSlideLinesArray.push(rtfLine)
                        }
                    })
                    let rtfSlideLines = rtfSlideLinesArray.join("");                    
                    let rtfSlide = presentationTemplates.slideText.replace("$TEXT_LINES", rtfSlideLines);
                    let pro7SlideString = presentationTemplates.slide.replace("$RTF_DATA", rtfSlide);
                    pro7SlideString = pro7SlideString.replace("\\\\par\\\\pard}", "}")
                    pro7SlideString = pro7SlideString.replace(/\$UUID/gm, function(){
                        return uuidv4()
                    });
                    pro7SlideString = pro7SlideString.replace("$SLIDE_UUID", slideId)
                    pro7SlidesArray.push(pro7SlideString)
                    slideIdentifierGuids.push(slideId);
                }
            })        

            var presentationString = presentationTemplates.presentation.replace("$PRESENTATION_NAME", "testinggg");

            pro7SlidesArray.forEach(function (value, i) {
                let slideIdentifier = presentationTemplates.slideIdentifier.replace("$SLIDE_UUID", slideIdentifierGuids[i])
                slideIdentifiers.push(slideIdentifier);
            })
            let slideIdentifiersString = slideIdentifiers.join("");
            presentationString = presentationString.replace("$SLIDE_IDENTIFIERS", slideIdentifiersString)
            presentationString = presentationString.replace("$SLIDES", pro7SlidesArray.join("\n"));
            presentationString = presentationString.replace(/\$UUID/gm, function(){
                        return uuidv4()
                    });
            fs.writeFileSync("./presentationData.txt", presentationString, err => {
                console.error(err);
            });
            context.log("Presentation data parsed into format successfully.")
              
            const command = path.resolve('./protoc/bin/protoc.exe');
            const args = [
            '--encode', 'rv.data.Presentation',
            './proto/presentation.proto',
            '--proto_path', './proto/'
            ];

            const result = child.spawnSync(command, args, {
            input: fs.readFileSync('./presentationData.txt'),
            stdio: ['pipe', 'pipe', 'pipe'],
            });
            if (result.error) {
                context.error('Error executing command:', result.error);
                process.exit(1);
            }
            fs.writeFileSync(`./${outputFilePath}`, result.stdout);
            context.log("File created successfully")          
        }
        context.log("Uploading file to SharePoint...")       
        const destinationFolder = ((jsonFileInfo.parentReference.path).split("root:/")[0]) + "root:/proPresenter Files"
        context.log("Destination path: " + destinationFolder)

        const folderInfo = await apiCall(`https://graph.microsoft.com/v1.0/${destinationFolder}`, {
            method: "GET",
            headers:{
                "Authorization": `Bearer ${accessToken}`,
                "Accept": "application/json"
            }
        })

        await apiCall(`https://graph.microsoft.com/v1.0/drives/${requestData.driveId}/items/${folderInfo.id}:/${outputFilePath}:/content`, {
            method: "PUT",
            headers: {
                "Content-Type": "text/plain",
                "Authorization": `Bearer ${accessToken}`
            },
            body: fs.readFileSync(`./${outputFilePath}`),
            redirect: "follow"
        })

        context.log("File uploaded successfully.")
        context.log("Deleting temporary files")
        context.log("Temporary Files deleted")
        return { body: `This worked!` };
    }
});
