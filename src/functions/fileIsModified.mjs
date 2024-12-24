import { app } from "@azure/functions";
import * as textract from '@markell12/textract'
import * as fs from 'fs';
import { writeFile } from 'node:fs/promises'
import { Readable } from 'node:stream'
import { v4 as uuidv4} from 'uuid';
import * as child from 'child_process'
import path from 'path';
import env from 'env-var';

//todo
//error handling
//code cleanup
//convert all API calls to promise-based
//error checking on all API calls
//sharepoint replacing file instead of creating new one
//variable cleanup

function b64(text){
    return Buffer.from(text).toString('base64');
}

async function extractTextFromPptx(filePath) {
    let pptxText;
    try {
        pptxText = await new Promise((resolve, reject) => {
            textract.fromFileWithPath(filePath, {
                "preserveLineBreaks":true,
                "preserveOnlyMultipleLineBreaks":false,
                "tesseract.lang":"rus"
            }, (error, text) => {
                if (error) {
                    reject(error);
                } else {
                    resolve(text);
                }
            });
        });
    } catch (error) {
        console.error("Error extracting text:", error);
    }
    return pptxText;
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

        var accessToken, fileInfoResponse;
        const requestBody = await request.text();
        const requestData = JSON.parse(requestBody)
        context.log("Received request: " + requestBody)
        context.log("Retrieving Access token...")
        await fetch(`https://login.microsoftonline.com/${config.tenantId}/oauth2/token`, {
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
            redirect: "follow"
        })
            .then((response) => response.text())
            .then((result) => {
                context.log("Retrieved token successfully.");
                const jsonTokenData = JSON.parse(result);
                accessToken = jsonTokenData.access_token;
            })
            .catch((error) => context.error(error));
        
        context.log("Retrieving file information...")
        await fetch(`https://graph.microsoft.com/v1.0/drives/${requestData.driveId}/items/${requestData.driveItemId}`,{
            method: "GET",
            headers:{
                "Authorization": `Bearer ${accessToken}`
            }
        })
            .then((response) => response.text())
            .then((result) => {
                context.log("File information retrieved successfully.")
                fileInfoResponse = result;
            })
            .catch((error) => context.error(error));

        const jsonFileInfo = JSON.parse(fileInfoResponse);
        const downloadUrl = jsonFileInfo["@microsoft.graph.downloadUrl"];
        context.log("Downloading file...");
        const response = await fetch(downloadUrl)
            .catch((error) => context.error(error));
        const stream = Readable.fromWeb(response.body);
        await writeFile('powerpoint.pptx', stream);
        context.log("File downloaded successfully.");
        
        context.log("Starting file conversion...")
        
        let pptxText = await extractTextFromPptx("./powerpoint.pptx");

        const textSlides = pptxText.split("\n\n");
        var outputFilePath
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
            fs.writeFileSync("./presentationString.txt", presentationString, err => {
                if (err) {
                    console.error(err);
                } else {
                    context.log("Presentation data parsed into format successfully.")
                }
            });
              
                // Define the command and arguments
                const command = path.resolve('./protoc/bin/protoc.exe'); // Adjust the path if necessary
                const args = [
                '--encode', 'rv.data.Presentation',
                './proto/presentation.proto',
                '--proto_path', './proto/'
                ];

                // Run the command
                const result = child.spawnSync(command, args, {
                input: fs.readFileSync('./presentationString.txt'), // Input redirection
                stdio: ['pipe', 'pipe', 'pipe'], // Configures stdin, stdout, and stderr
                });
                if (result.error) {
                    context.error('Error executing command:', result.error);
                    process.exit(1);
                }
                fs.writeFileSync(`./${outputFilePath}`, result.stdout);            
        }
        context.log("Uploading file to SharePoint...")       
        const destinationFolder = ((jsonFileInfo.parentReference.path).split("root:/")[0]) + "root:/proPresenter Files"
        var fileName = outputFilePath;
        var destinationPath = destinationFolder + "/" + outputFilePath;
        console.log(destinationPath)
        context.log("Destination path: " + destinationPath)
        var isNameUnique = false;
        var i = 1;
        while (!isNameUnique) {
            destinationPath = destinationFolder + fileName;
            await fetch(`https://graph.microsoft.com/v1.0/${destinationPath}`,{ // check if file exists
                method: "GET",
                headers:{
                    "Authorization": `Bearer ${accessToken}`,
                    "Accept": "application/json"
                }
            })
                .then((response) => response.text())
                .then((result) => {
                    context.log(result)
                    if(result == '{"error":{"code":"itemNotFound","message":"The resource could not be found."}}'){
                        context.log("File name is unique. Uploading the file...")
                        isNameUnique = true;

                        const file = fs.readFileSync(`./${outputFilePath}`);

                        const requestOptions = {
                        method: "PUT",
                        headers: {
                            "Content-Type": "text/plain",
                            "Authorization": `Bearer ${accessToken}`
                        },
                        body: file,
                        redirect: "follow"
                        };

                        fetch(`https://graph.microsoft.com/v1.0/drives/b!4UHkXyJCHU-eOG0diOb45t0ezJHvmUFHgfS4Dq_i6-rCrjwkY5MaQYaarmbje6No/items/01STCM33YWJBC2LLXEIZBIN346XOQGEGDL:/${outputFilePath}:/content`, requestOptions)
                        .then((response) => response.text())
                        .then((result) => {
                            context.log(result)
                            context.log("File uploaded successfully.")
                            context.log("Deleting temporary files")

                            fs.unlinkSync("./presentationString.txt");
                            fs.unlinkSync("./powerpoint.pptx");
                            fs.unlinkSync(`./${outputFilePath}`);
                            context.log("Temporary Files deleted")

                        })
                        .catch((error) => context.error(error));
                    }else{
                        context.log("File name is not unique.")
                        if(i == 1){
                            fileName += fileName + `( ${i})`
                        }else{
                            if(config.proPresenterVersion == 6){
                                fileName = fileName.replace(/ (.*?).pro6/, ` (${i}).pro6`)
                            } else if (config.proPresenterVersion == 7){
                                fileName = fileName.replace(/ (.*?).pro/, ` (${i}).pro`)
                            }
                        }
                        i++;
                    }
                })
                .catch((error) => context.error(error));
        }
        return { body: `This worked!` };
    }
});
