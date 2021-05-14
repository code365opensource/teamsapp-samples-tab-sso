import { useEffect, useState } from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { Button, Flex, Segment, Provider, teamsTheme, Text } from "@fluentui/react-northstar";
import { BrowserRouter as Router, Route } from "react-router-dom";
import jwtdecode from "jwt-decode";
import crypto from "crypto";
import * as graph from "@microsoft/microsoft-graph-client";


function AuthStart() {
  useEffect(() => {
    microsoftTeams.initialize();
    microsoftTeams.getContext((context: microsoftTeams.Context) => {

      let tenant = context['tid'];
      let client_id = "692fb9c1-02ab-4bc0-bfaf-c270cedf85b8";
      let queryParams: any = {
        tenant: `${tenant}`,
        client_id: `${client_id}`,
        response_type: "token", //token_id in other samples is only needed if using open ID
        redirect_uri: window.location.origin + "/auth-end",
        scope: "https://graph.microsoft.com/User.Read",
        nonce: crypto.randomBytes(16).toString('base64')
      }

      let url = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize?`;
      queryParams = new URLSearchParams(queryParams).toString();
      let authorizeEndpoint = url + queryParams;
      window.location.assign(authorizeEndpoint);

    });
  }, [])
  return <Text content="身份认证开始..."></Text>
}

function AuthEnd() {
  const getHashParameters = () => {
    let hashParams: any = {};
    window.location.hash.substr(1).split("&").forEach(function (item) {
      let [key, value] = item.split('=');
      hashParams[key] = decodeURIComponent(value);
    });
    return hashParams;
  }

  useEffect(() => {
    microsoftTeams.initialize();

    //The Azure implicit grant flow injects the result into the window.location.hash object. Parse it to find the results.
    let hashParams = getHashParameters();

    //If consent has been successfully granted, the Graph access token should be present as a field in the dictionary.
    if (hashParams["access_token"]) {
      //Notifify the showConsentDialogue function in Tab.js that authorization succeeded. The success callback should fire. 
      microsoftTeams.authentication.notifySuccess(hashParams["access_token"]);
    } else {
      microsoftTeams.authentication.notifyFailure("Consent failed");
    }
  }, [])

  return <Text content="身份认证结束..."></Text>

}

function Home() {

  const [authToken, setAuthToken] = useState<string>();
  const [graphToken, setGraphToken] = useState<string>();
  const [userName, setUserName] = useState<string>();
  const [fileContent, setFileContent] = useState<string>();

  return (
    <Flex column fill gap="gap.medium">
      <Button content="获取Auth Token" onClick={() => {
        microsoftTeams.authentication.getAuthToken({
          successCallback: (token: string) => {
            setAuthToken(token);
          }
        })
      }}></Button>
      <Segment content={authToken} color="red"></Segment>
      <Button content="获取Graph Token" onClick={async () => {
        let serverURL = `api/token?ssoToken=${authToken}`;
        let response = await fetch(serverURL);
        if (response) {
          let data = await response.json();
          if (!response.ok && data.error === '要求授权') {
            microsoftTeams.authentication.authenticate({
              url: window.location.origin + "/auth-start",
              width: 600,
              height: 535,
              successCallback: (result) => { setGraphToken(result); },
              failureCallback: (reason) => { }
            });

          } else if (!response.ok) {
            console.log(data.error);

          } else {
            setGraphToken(data["access_token"]);
          }
        }
      }}></Button>
      <Segment content={graphToken} color="blue"></Segment>
      <Button content="获取用户名" onClick={async () => {
        if (authToken && !graphToken) {
          setUserName("本地读取到的用户名:" + (jwtdecode(authToken) as any).name);
        }
        else if (graphToken) {
          // let graphmeEndpoint = `https://graph.microsoft.com/v1.0/me`;
          // let graphRequestParams = {
          //   method: 'GET',
          //   headers: {
          //     "authorization": "bearer " + graphToken
          //   }
          // }

          // let response = await fetch(graphmeEndpoint, graphRequestParams);
          // if (response) {
          //   if (!response.ok) {
          //     console.error("出现错误: ", response);
          //   }
          //   else {
          //     const user = await response.json();
          //     setUserName(user.DisplayName);
          //   }
          // }

          const client = graph.Client.init({
            authProvider: (done: any) => {
              done(null, graphToken);
            }
          })

          const user = await client.api("/me").get();
          console.log(user);
          setUserName(user.displayName);

        }
      }}></Button>
      <Segment content={userName} color="green"></Segment>
      <Button content="获取文件内容"></Button>
      <Segment content={fileContent} color="black"></Segment>
    </Flex>

  );
}


function App() {
  //Teams客户端环境初始化
  microsoftTeams.initialize();

  return <Provider theme={teamsTheme}>
    <Router>
      <Route path="/" component={Home} exact></Route>
      <Route path="/auth-start" component={AuthStart} exact></Route>
      <Route path="/auth-end" component={AuthEnd} exact></Route>
    </Router>
  </Provider>

}

export default App;
