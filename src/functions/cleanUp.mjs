import * as fs from 'fs';

export function cleanUp(context){
    context.log("Deleting temporary files")
    fs.unlinkSync("./presentationData.txt", (err) => {
        context.error("Error deleting files: " + err)
    })
    context.log("Temporary Files deleted")
}