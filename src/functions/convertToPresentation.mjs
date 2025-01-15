import env from 'env-var';
import * as fs from 'fs/promises';
import path from 'path';
import * as child from 'child_process'
import { v4 as uuidv4} from 'uuid';

export async function convertToPresentation(powerPoint, context){
    var presentationFilePath;
    const proPresenterVersion = env.get("PRO_PRESENTER_VERSION").required().asIntPositive();
    context.log("Retrieved ProPresenter version: " + proPresenterVersion)
    if(proPresenterVersion == 6){ //untested
        const presentationTemplates = {
            presentationHeader: fs.readFileSync('./pro6Templates/presentationHeader.txt').toString(),
            presentationFooter: fs.readFileSync('./pro6Templates/presentationFooter.txt').toString(),
            slide: fs.readFileSync('./pro6Templates/presentationSlide.txt').toString()
        }
        presentationFilePath = "C:/local/temp/presentation.pro6"
        var pro6SlidesArray = [];
        (powerPoint.slides).forEach(slide => {
            let rtfSlideLinesArray = [];
            slide.forEach(line =>{
                let rtfLine = ("{\\fs200\\outl0\\strokewidth-20\\strokec1\\f0 {\\cf3\\ltrch " + line + "}\\li0\\sa0\\sb0\\fi0\\qc\\par}")
                rtfSlideLinesArray.push(rtfLine)
            })
            let rtfSlideLines = rtfSlideLinesArray.join("\\n")
            let rtfSlide = ("{\\rtf1\\prortf1\\ansi\\ansicpg1252\\uc1\\htmautsp\\deff2{\\fonttbl{\\f0\\fcharset0 Times New Roman;}}{\\colortbl;\\red0\\green0\\blue0;\\red255\\green255\\blue255;\\red250\\green235\\blue215;}\\loch\\hich\\dbch\\pard\\slleading0\\plain\\ltrpar\\itap0{\\lang1033\\fs100\\outl0\\strokewidth-20\\strokec1\\f2\\cf1 \\cf1\\qc \n" + rtfSlideLines + "\n } \n }");
            let b64RTFText = b64(rtfSlide)
            let pro6SlideString = presentationTemplates.slide;
            pro6SlideString = pro6SlideString.replace("$PLAIN_TEXT", b64(slide.join("\n")));
            pro6SlideString = pro6SlideString.replace("$RTF_TEXT", b64RTFText);
            pro6SlideString = pro6SlideString.replace("$SLIDE_UUID", uuidv4());
            pro6SlideString = pro6SlideString.replace("$TEXTBOX_UUID", uuidv4());
            pro6SlidesArray.push(pro6SlideString)
        })        

        const presentationString = presentationTemplates.presentationHeader + pro6SlidesArray.join() + presentationTemplates.presentationFooter;
        fs.writeFileSync(presentationFilePath, presentationString, err => {
            if (err) {
                console.error(err);
            } else {
                context.log("pro6 File created successfully.")
            }
        });
        
    } else if (proPresenterVersion == 7){
        try{
            const presentationTemplates = {
                presentation: await fs.readFile('./pro7Templates/presentation.txt').toString(),
                slide: await fs.readFile('./pro7Templates/slide.txt').toString(),
                slideText: await fs.readFile('./pro7Templates/slideText.txt').toString(),
                textLine: await fs.readFile('./pro7Templates/textLine.txt').toString(),
                slideIdentifier: await fs.readFile('./pro7Templates/slideIdentifier.txt').toString()
            }
            context.log("Read template files.")
            var pro7SlidesArray = [],
                slideIdentifierGuids = [],
                slideIdentifiers = [];
            presentationFilePath = "C:/local/temp/presentation.pro";
            
            (powerPoint.slides).forEach(slide => {
                let slideId = uuidv4();
                let rtfSlideLinesArray = [];
                slide.forEach(line =>{
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
            })        
            context.log("Created slides")
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

            context.log("Created presentation string")
            await fs.writeFile("c:/local/temp/presentationData.txt", presentationString, err => {
                console.error(err);
            });
            context.log("Presentation data parsed into format successfully.")
              
            const command = path.resolve('./protoc/src/protoc.exe');
            const args = [
            '--encode', 'rv.data.Presentation',
            './proto/presentation.proto',
            '--proto_path', './proto/'
            ];
    
            const result = child.spawn(command, args, {
            input: await fs.readFile('c:/local/temp/presentationData.txt'),
            stdio: ['pipe', 'pipe', 'pipe'],
            });
            if (result.error) {
                context.error('Error executing command:', result.error);
                process.exit(1);
            }
            fs.writeFile(presentationFilePath, result.stdout);
            context.log("File created successfully")     
            return presentationFilePath;
        }
        catch(err){
            context.error(err)
        } 
    }else{
        context.error("Invalid ProPresenter version")
    }
}