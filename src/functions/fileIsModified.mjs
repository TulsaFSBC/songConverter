import { app } from "@azure/functions";
import { getTextExtractor } from "office-text-extractor";
import * as fs from 'fs';
import { writeFile } from 'node:fs/promises'
import { Readable } from 'node:stream'
import Format from '../../node_modules/rtf/lib/format.js'
import Colors from '../../node_modules/rtf/lib/colors.js'
import Fonts from '../../node_modules/rtf/lib/fonts.js'
import RTF from '../../node_modules/rtf/lib/rtf.js'
import { v4 as uuidv4} from 'uuid';

function rtfText(plainText) {
    let rtfContent = new RTF();
    let textFormat = new Format();
    textFormat.fontSize = 100;
    textFormat.color = Colors.WHITE;
    textFormat.font = Fonts.TIMES_NEW_ROMAN
    if(plainText.length > 0){
        rtfContent.writeText(plainText, textFormat);
        let outputRTF;
        rtfContent.createDocument(
            function(err,output){
                outputRTF = output
            }
        )
        return outputRTF
    } else {
        console.error("plainText property is empty. Please set before running this function.")
        return null
    }  
}

function b64(text){
    return Buffer.from(text).toString('base64');
}

app.http('fileIsModified', {
    methods: ['GET','POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        const extractor = getTextExtractor()
        var accessToken, fileInfoResponse;
        const requestBody = await request.text();
        const requestData = JSON.parse(requestBody)
        context.log("Received request: " + requestBody)
        /*context.log("Retrieving Access token...")
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
        context.log("File downloaded successfully.");*/
        
        context.log("Starting file conversion...")
        const presentationHeader = fs.readFileSync('./pro6Templates/presentationHeader.txt').toString(),
              presentationFooter = fs.readFileSync('./pro6Templates/presentationFooter.txt').toString(),
              slideTemplate = fs.readFileSync('./pro6Templates/presentationSlide.txt').toString();

        const text = await extractor.extractText({ input: "./powerpoint.pptx", type: 'file' })
        const textSlides = text.split("---");
        
        var pro6SlidesArray = []
        textSlides.forEach(slide => {
            let b64PlainText = b64(slide);
            let slideLines = slide.split("\n");
            let rtfSlideLinesArray = [];
            slideLines.forEach(line =>{
                if (line != ""){
                    let rtfLine = ("{\\fs200\\outl0\\strokewidth-20\\strokec1\\f0 {\\cf3\\ltrch " + line + "}\\li0\\sa0\\sb0\\fi0\\qc\\par}")
                    console.log("line: " + line)
                    console.log("rtf line: " + rtfLine)
                    rtfSlideLinesArray.push(rtfLine)
                }
            })
            let rtfSlideLines = rtfSlideLinesArray.join("\\n")
            let rtfSlide = ("{\\rtf1\\prortf1\\ansi\\ansicpg1252\\uc1\\htmautsp\\deff2{\\fonttbl{\\f0\\fcharset0 Times New Roman;}}{\\colortbl;\\red0\\green0\\blue0;\\red255\\green255\\blue255;\\red250\\green235\\blue215;}\\loch\\hich\\dbch\\pard\\slleading0\\plain\\ltrpar\\itap0{\\lang1033\\fs100\\outl0\\strokewidth-20\\strokec1\\f2\\cf1 \\cf1\\qc \n" + rtfSlideLines + "\n } \n }");
            //console.log(rtfSlide)
            let b64RTFText = b64(rtfSlide)
            let pro6SlideString = slideTemplate;
            pro6SlideString = pro6SlideString.replace("$PLAIN_TEXT", b64PlainText); //plugs in actual content into slide template
            pro6SlideString = pro6SlideString.replace("$RTF_TEXT", b64RTFText);
            pro6SlideString = pro6SlideString.replace("$SLIDE_UUID", uuidv4());
            pro6SlideString = pro6SlideString.replace("$TEXTBOX_UUID", uuidv4());
            pro6SlidesArray.push(pro6SlideString)
        })        

        const presentationString = presentationHeader + pro6SlidesArray.join() + presentationFooter;

        fs.writeFile('./pro6.pro6', presentationString, err => {
            if (err) {
                console.error(err);
            } else {
                context.log("pro6 File created successfully.")
            }
        });
        return { body: `This worked!` };
    }
});
