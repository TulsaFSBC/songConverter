import { app } from "@azure/functions";
import { getTextExtractor } from "office-text-extractor";
import * as fs from 'fs';
import { writeFile } from 'node:fs/promises'
import { Readable } from 'node:stream'

app.http('fileIsModified', {
    methods: ['GET','POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        const extractor = getTextExtractor()
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
        context.log("File downloaded successfully.");
        const stream = Readable.fromWeb(response.body);
        await writeFile('powerpoint.pptx', stream);
        context.log("File saved successfully.");
        
        context.log("Starting file conversion...")

        const rtfHeader = `{\\rtf1\\prortf1\\ansi\\ansicpg1252\\uc1\\htmautsp\\deff2` +
        `{\\fonttbl{\\f0\\fcharset0 Times New Roman;}{\\f2\\fcharset0 Georgia;}{\\f3\\fcharset0 Segoe UI;}}` +
        `{\\colortbl;\\red0\\green0\\blue0;\\red255\\green255\\blue255;\\red250\\green235\\blue215;}` +
        `\\loch\\hich\\dbch\\pard\\slleading0\\plain\\ltrpar\\itap0` +
        `{\\lang1033\\fs32\\outl0\\strokewidth-60\\strokec1\\f2\\cf1 \\cf1\\qc`;
        const rtfFooter = '\\li0\\sa0\\sb0\\fi0\\qc}}';


        const text = await extractor.extractText({ input: "./powerpoint.pptx", type: 'file' })
        const textSlides = text.split("---");

        var b64Slides = [];
        textSlides.forEach(slide => {
            let b64Slide = Buffer.from(slide).toString('base64');
            b64Slides.push(b64Slide);
        })
        //commnet
        var b64RTFSlides = [];
        textSlides.forEach(slide =>{
            let rtfLinesArray = [];
            let lines = slide.split("\n")
            lines.forEach(line =>{
                const unicodeText = line.split('').map(char => {
                    let code = char.charCodeAt(0);
                    return `\\u${code}?`;
                }).join('');
                rtfLinesArray.push(`{\\fs180\\outl0\\strokewidth-60\\strokec1\\f3{\\cf3\\ltrch ${unicodeText}}\\li0\\sa0\\sb0\\fi0\\qc\\par}`)
            })
            let rtfLines = rtfLinesArray.join('');
            let rtfContent = rtfHeader + rtfLines + rtfFooter;
            let b64RTFContent = Buffer.from(rtfContent).toString('base64');
            b64RTFSlides.push(b64RTFContent);
        })

        const slideHeader = fs.readFileSync('./pro6Templates/presentationHeader.txt').toString();
        const slideFooter = fs.readFileSync('./pro6Templates/presentationFooter.txt').toString();
        const slideTemplate = fs.readFileSync('./pro6Templates/presentationSlide.txt').toString();
        var pro6Slides = [];
        var i = 0;

        textSlides.forEach ( slide => {
            let plainTextReplaced = slideTemplate.replace('<NSString rvXMLIvarName="PlainText"></NSString>', '<NSString rvXMLIvarName="PlainText">' + b64Slides[i] + '</NSString>')
            let RTFReplaced = plainTextReplaced.replace('<NSString rvXMLIvarName="RTFData"></NSString>', '<NSString rvXMLIvarName="RTFData">' + b64RTFSlides[i] + '</NSString>')
            pro6Slides.push(RTFReplaced)
            i++;
        })

        const slides = pro6Slides.join('');
        const presentationString = slideHeader + slides + slideFooter;

        fs.writeFile('./pro6.pro6', presentationString, err => {
            if (err) {
                console.error(err);
            } else {
                context.log("pro6 File created successfully.")
            }
        });
        return { body: `Hello,` };
    }
});
