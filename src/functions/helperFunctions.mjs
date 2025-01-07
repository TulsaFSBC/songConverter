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
        return await JSON.parse(text); // Return the text for potential future use
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

/*const requestBody = await request.text()
    await context.log("Received request: " + await requestBody)
    return await JSON.parse(requestBody) */
