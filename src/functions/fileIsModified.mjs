import { app } from "@azure/functions";
import * as textract from '@markell12/textract'
import * as fs from 'fs';
import { writeFile } from 'node:fs/promises'
import { Readable } from 'node:stream'
import { v4 as uuidv4} from 'uuid';
import * as child from 'child_process'

//todo
//slide creation cleanup
//remove rtf module
//create correct presentation string
//upload to SP
//delete temp files after function executes: powerpoint.pptx, presentationString.txt, .pro file.
//Logging to a persistent source
//error handling
//config
//code cleanup

function b64(text){
    return Buffer.from(text).toString('base64');
}

app.http('fileIsModified', {
    methods: ['GET','POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        const presentationVersion = 7;
        var accessToken, fileInfoResponse;
        const requestBody = await request.text();
        const requestData = JSON.parse(requestBody)
        context.log("Received request: " + requestBody)
        context.log("Retrieving Access token...")
        await fetch(`https://login.microsoftonline.com/${process.env.tenant_id}/oauth2/token`, {
            method: "POST",
            headers: {
                "content-type": "application/x-www-form-urlencoded"
            },
            body: new URLSearchParams({
                "grant_type": "client_credentials",
                "client_id": process.env.client_id,
                "client_secret": process.env.client_secret,
                "resource": "https://graph.microsoft.com"
            }),
            redirect: "follow"
        })
            .then((response) => response.text())
            .then((result) => {
                console.log("Retrieved token successfully.");
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
                console.log("File information retrieved successfully.")
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
        const pro6PresentationHeader = fs.readFileSync('./pro6Templates/presentationHeader.txt').toString(),
              pro6PresentationFooter = fs.readFileSync('./pro6Templates/presentationFooter.txt').toString(),
              pro6SlideTemplate = fs.readFileSync('./pro6Templates/presentationSlide.txt').toString();

        const pro7PresentationTemplate = fs.readFileSync('./pro7Templates/presentation.txt').toString(),
              pro7SlideTemplate = fs.readFileSync('./pro7Templates/slide.txt').toString(),
              pro7SlideTextTemplate = fs.readFileSync('./pro7Templates/slideText.txt').toString(),
              pro7TextLineTemplate = fs.readFileSync('./pro7Templates/textLinee.txt').toString(),
              pro7SlideIdentifierTemplate = fs.readFileSync('./pro7Templates/slideIdentifier.txt').toString()

        let config = {
            "preserveLineBreaks":true,
            "preserveOnlyMultipleLineBreaks":false,
            "tesseract.lang":"rus"
        }
        let pptxText = "";
        async function extractTextFromPptx() {
            try {
                pptxText = await new Promise((resolve, reject) => {
                    textract.fromFileWithPath("./powerpoint1.pptx", config, (error, text) => {
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
        }
        await extractTextFromPptx();
        const textSlides = pptxText.split("\n\n");
        var outputFilePath
        if(presentationVersion == 6){
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
                    let pro6SlideString = pro6SlideTemplate;
                    pro6SlideString = pro6SlideString.replace("$PLAIN_TEXT", b64PlainText); //plugs in actual content into slide template
                    pro6SlideString = pro6SlideString.replace("$RTF_TEXT", b64RTFText);
                    pro6SlideString = pro6SlideString.replace("$SLIDE_UUID", uuidv4());
                    pro6SlideString = pro6SlideString.replace("$TEXTBOX_UUID", uuidv4());
                    pro6SlidesArray.push(pro6SlideString)
                }
            })        

            const presentationString = pro6PresentationHeader + pro6SlidesArray.join() + pro6PresentationFooter;
            fs.writeFileSync(`./${outputFilePath}`, presentationString, err => {
                if (err) {
                    console.error(err);
                } else {
                    context.log("pro6 File created successfully.")
                }
            });
        } else if (presentationVersion == 7){
            var pro7SlidesArray = [],
                slideIdentifierGuids = [],
                slideIdentifiers = [];
            outputFilePath = `${(jsonFileInfo.name).replace(".pptx", ".pro")}`
            
            textSlides.forEach(slide => {
                if(slide != ""){
                    let slideId = uuidv4()
                    let slideLines = slide.split("\n");
                    let rtfSlideLinesArray = [];
                    slideLines.forEach(line =>{
                        if (line != ""){
                            let rtfLine = pro7TextLineTemplate.replace("$TEXT", line)
                            rtfSlideLinesArray.push(rtfLine)
                        }
                    })
                    slideLines[slideLines.length - 1].replace(/\\par\\pard/gm, "")
                    let rtfSlideLines = rtfSlideLinesArray.join("");
                    console.log(rtfSlideLines)                    
                    let rtfSlide = pro7SlideTextTemplate.replace("$TEXT_LINES", rtfSlideLines);
                    rtfSlide = rtfSlide.replace("$SLIDE_UUID", slideId)
                    let pro7SlideString = pro7SlideTemplate.replace("$RTF_DATA", rtfSlide);
                    pro7SlideString = pro7SlideString.replace(/\$UUID/gm, function(){
                        return uuidv4()
                    });
                    pro7SlidesArray.push(pro7SlideString)
                    slideIdentifierGuids.push(slideId);
                }
            })        

            var presentationString = pro7PresentationTemplate.replace("$PRESENTATION_NAME", "testinggg");

            pro7SlidesArray.forEach(function (value, i) {
                let slideIdentifier = pro7SlideIdentifierTemplate.replace("$SLIDE_UUID", slideIdentifierGuids[i])
                slideIdentifiers.push(slideIdentifier);
            })
            let slideIdentifiersString = slideIdentifiers.join("");
            presentationString = presentationString.replace("$SLIDE_IDENTIFIERS"), slideIdentifiersString
            presentationString = presentationString.replace("$SLIDES", pro7SlidesArray.join("\n"));
            presentationString = presentationString.replace(/\$UUID/gm, function(){
                        return uuidv4()
                    });
            //presentationString = presentationString.replaceAll("\\", "\\\\")
            fs.writeFileSync(`./presentationString1.txt`, presentationString, err => {
                if (err) {
                    console.error(err);
                } else {
                    context.log("pro File created successfully.")
                }
            });
            

            /*await child.exec(`".\\protoc\\bin\\protoc.exe" --encode="rv.data.Presentation" --proto_path ".\\proto" ".\\proto\\propresenter.proto" < "${presentationString}" > ".\\${outputFile}"`, (err, stdout, stderr) => {
                if (err) {
                  // node couldn't execute the command
                  return;
                }
              
                // the *entire* stdout and stderr (buffered)
                console.log("I've been here!")
                console.log(`stdout: ${stdout}`);
                console.log(`stderr: ${stderr}`);

              }); */
              
                const outputFile = fs.createWriteStream(`./${outputFilePath}`);
                const inputFile = fs.createReadStream("./presentationString.txt")
                const protoc = child.spawn('.\\protoc\\bin\\protoc.exe', [
                '--encode=rv.data.Presentation',
                '--proto_path=./proto',
                './proto/propresenter.proto'
                ]);

                // Pipe input and output
                inputFile.pipe(protoc.stdin);
                protoc.stdout.pipe(outputFile);

                protoc.on('close', (code) => {
                    console.log(`Process exited with code: ${code}`);
                });

                protoc.stderr.on('data', (data) => {
                    console.error(`stderr: ${data}`);
                });

        }
        
/*
        context.log("Uploading file to SharePoint...")
        const destinationFolder = ((jsonFileInfo.parentReference.path).split("root:/")[0]) + "root:/proPresenter Files"
        var fileName = outputFile;
        var destinationPath;
        context.log("Destination path: " + destinationPath)
        var isNameUnique = false;
        var i = 1;
        while (!isNameUnique) {
            destinationPath = destinationFolder //+ fileName;
            await fetch(`https://graph.microsoft.com/v1.0/${destinationPath}`,{ // check if file exists
                method: "GET",
                headers:{
                    "Authorization": `Bearer ${accessToken}`,
                    "Accept": "application/json"
                }
            })
                .then((response) => response.text())
                .then((result) => {
                    console.log(result)
                    if(result == '{"error":{"code":"itemNotFound","message":"The resource could not be found."}}'){
                        context.log("File name is unique. Uploading the file...")
                        isNameUnique = true;
                    }else{
                        context.log("File name is not unique.")
                        if(i == 1){
                            fileName += fileName + `( ${i})`
                        }else{
                            if(presentationVersion == 6){
                                fileName = fileName.replace(/ (.*?).pro6/, ` (${i}).pro6`)
                            } else if (presentationVersion == 7){
                                fileName = fileName.replace(/ (.*?).pro/, ` (${i}).pro`)
                            }
                            
                        }
                        i++;
                    }
                })
                .catch((error) => context.error(error));
        }
        */
        return { body: `This worked!` };
    }
});
