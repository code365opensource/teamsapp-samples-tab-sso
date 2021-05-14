import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import jwt_decode, { JwtPayload } from "jwt-decode";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    const ssoToken = req.query.ssoToken as string;
    let tenantId = jwt_decode<JwtPayload>(ssoToken)['tid']; //获取租户的编号
    let accessTokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    let accessTokenQueryParams = {
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        client_id: "692fb9c1-02ab-4bc0-bfaf-c270cedf85b8",
        client_secret: "W86QwNW8g~CP.v9dA-bKj_OuM~iJJ75yd3",
        assertion: req.query.ssoToken as string,
        scope: 'https://graph.microsoft.com/.default',
        requested_token_use: "on_behalf_of",
    };

    let body = new URLSearchParams(accessTokenQueryParams).toString();

    let accessTokenReqOptions = {
        method: 'POST',
        headers: {
            Accept: "application/json",
            "Content-Type": "application/x-www-form-urlencoded"
        },
        body: body
    };

    const fetch = require("node-fetch");

    let response = await fetch(accessTokenEndpoint, accessTokenReqOptions);
    let data = await response.json();
    if (!response.ok) {
        if ((data.error === 'invalid_grant') || (data.error === 'interaction_required')) {
            context.res = {
                status: "403",
                body: {
                    error: "要求授权"
                }
            }
        } else {
            context.res = {
                status: "500",
                body: {
                    error: "未知错误：无法交换令牌"
                }
            }
        }
    } else {
        context.res = {
            body: data
        }
    }

};

export default httpTrigger;