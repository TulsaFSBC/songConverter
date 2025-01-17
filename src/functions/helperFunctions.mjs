export function b64(text){
    return Buffer.from(text).toString('base64');
}

export async function apiCall(url, requestOptions){
    const response = await fetch(url, requestOptions);
    const jsonData = await response.text();
    const data = await JSON.parse(jsonData);
    const statusCode = response.status
    return {data, statusCode};
}

export async function receiveRequest(request, context) {
    try {
        const text = await request.text();
        context.log("Request Received: " + text);
        if(typeof text == "string"){
            return await JSON.parse(text);
        } else {
            return text;
        }
    } catch (error) {
        context.error("Error receiving request: " + error);
        throw error;
    }
}

export function sleep(ms) {
    return new Promise((resolve) => {
      setTimeout(resolve, ms);
    });
  }
