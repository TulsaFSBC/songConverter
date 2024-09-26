const { app } = require('@azure/functions');
import { getTextExtractor } from 'office-text-extractor'
import fs from 'fs';

app.http('fileIsCreated', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log(`Http function processed request for url "${request.url}"`);
        const name = request.query.get('name') || await request.text() || 'world';
//my script
            const extractor = getTextExtractor()
            //const path = './ppt.pptx' REPLACE WITH FILE SUBMITTED BY PA FLOW
            const text = await extractor.extractText({ input: path, type: 'file' })
            const textSlides = text.split("---");

            var b64Slides = [];
            textSlides.forEach(slide => {
                let b64Slide = Buffer.from(slide).toString('base64');
                b64Slides.push(b64Slide);
            })

            var b64RTFSlides = [];
            const rtfHeader = `{\\rtf1\\prortf1\\ansi\\ansicpg1252\\uc1\\htmautsp\\deff2` +
            `{\\fonttbl{\\f0\\fcharset0 Times New Roman;}{\\f2\\fcharset0 Georgia;}{\\f3\\fcharset0 Segoe UI;}}` +
            `{\\colortbl;\\red0\\green0\\blue0;\\red255\\green255\\blue255;\\red250\\green235\\blue215;}` +
            `\\loch\\hich\\dbch\\pard\\slleading0\\plain\\ltrpar\\itap0` +
            `{\\lang1033\\fs32\\outl0\\strokewidth-60\\strokec1\\f2\\cf1 \\cf1\\qc`;

            const rtfFooter = '\\li0\\sa0\\sb0\\fi0\\qc}}';

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
            const slideHeader = fs.readFileSync('./presentationSrc/presentationHeader.txt').toString();
            const slideFooter = fs.readFileSync('./presentationSrc/presentationFooter.txt').toString();
            const slideTemplate = fs.readFileSync('./presentationSrc/presentationSlide.txt').toString();
            var pro6Slides = [];
            var i = 0;

            textSlides.forEach(slide =>{
                let plainTextReplaced = slideTemplate.replace('<NSString rvXMLIvarName="PlainText"></NSString>', '<NSString rvXMLIvarName="PlainText">' + b64Slides[i] + '</NSString>')
                let RTFReplaced = plainTextReplaced.replace('<NSString rvXMLIvarName="RTFData"></NSString>', '<NSString rvXMLIvarName="RTFData">' + b64RTFSlides[i] + '</NSString>')
                pro6Slides.push(RTFReplaced)
                i+= 1;
            })

            const slides = pro6Slides.join('');

            const presentationString = slideHeader + slides + slideFooter;

            fs.writeFile('./pro6.pro6', presentationString, err => {
                if (err) {
                console.error(err);
                } else {
                // file written successfully
                //upload file to needed sharepoint library
                }
            });

//end my script
        return { body: `Hello, ${name}!` };
    }
});
