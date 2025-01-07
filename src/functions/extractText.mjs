import { getTextExtractor } from 'office-text-extractor'

export async function extractText(fileBuffer, context){
    context.log("Extracting text from file...")
    const text = await getTextExtractor().extractText({ input: fileBuffer, type: 'buffer' })
    let powerPoint = {
        slides: []
    }
    let slidesArray = text.split("---")
    slidesArray.forEach(slide => {
        let textLines = slide.split("\n");
        textLines.forEach((line, index) => {
            if(line === ""){
                textLines.splice(index, 1);
            }
        })
        powerPoint.slides.push(textLines);
    })

    if(typeof powerPoint.slides[0] != undefined){
        context.log("Text extracted from file successfully.")
    }else{
        context.error("Error extracting text from file.")
    }
    return powerPoint;
}