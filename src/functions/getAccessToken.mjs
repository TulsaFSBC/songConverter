import env from 'env-var';
import {apiCall} from './helperFunctions.mjs';
import {ClientSecretCredential} from "@azure/identity";
import { SecretClient } from "@azure/keyvault-secrets";

export async function getAccessToken(context) {
    const environment = {
        tenantId: env.get("TENANT_ID").required().asString(),
        clientId: env.get("CLIENT_ID").required().asString(),
        clientSecret: env.get("CLIENT_SECRET").required().asString(),
        keyVaultURL: env.get("KEY_VAULT_URL").required().asUrlString()
    };

    const client = new SecretClient(
        environment.keyVaultURL,
        new ClientSecretCredential(
            environment.tenantId,
            environment.clientId,
            environment.clientSecret
        )
    );

    context.log("Retrieving current access token...");
    try {
        const secretFromVault = await client.getSecret("accessToken");
        context.log("Retrieval successful.\nValidating token...");

        if (secretFromVault.properties.expiresOn > Date.now()) {
            context.log("Access token is valid.");
            return secretFromVault.value;
        }

        context.log("Access token invalid... \nRetrieving new access token");
        const accessTokenResponse = await apiCall(
            `https://login.microsoftonline.com/${environment.tenantId}/oauth2/token`,
            {
                method: "POST",
                headers: {
                    "content-type": "application/x-www-form-urlencoded"
                },
                body: new URLSearchParams({
                    "grant_type": "client_credentials",
                    "client_id": environment.clientId,
                    "client_secret": environment.clientSecret,
                    "resource": "https://graph.microsoft.com"
                }),
                redirect: "follow"
            }
        );

        const accessToken = accessTokenResponse.data.access_token;
        if (accessToken) {
            context.log("New access token retrieved. \nSaving access token to Azure Key Vault...");
            await client.setSecret("accessToken", accessToken, {
                expiresOn: new Date(accessTokenResponse.data.expires_on * 1000)
            });
            return accessToken;
        }

        throw new Error("Failed to retrieve new access token. Response: " + JSON.stringify(accessTokenResponse));
    } catch (error) {
        context.error("Error in getAccessToken: " + error);
        throw error;
    }
}