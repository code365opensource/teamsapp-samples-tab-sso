import { useEffect, useState } from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { Button, Flex, Segment, Provider, teamsTheme, Text } from "@fluentui/react-northstar";
import { BrowserRouter as Router, Route } from "react-router-dom";
import jwtdecode from "jwt-decode";
import crypto from "crypto";
import * as graph from "@microsoft/microsoft-graph-client";

const scope = "https://graph.microsoft.com/User.Read https://graph.microsoft.com/Files.Read";


// 这个组件用来在弹出的一个对话框中，去请求身份验证，这里会跳到Azure的登陆页面，并且在成功后跳回到对应的页面（用来接收access token），这里用到的是典型的 code-grant 的授权流。
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
        scope: scope,
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
// 这个组件用来接收身份验证的结果
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
    let hashParams = getHashParameters();
    if (hashParams["access_token"]) {
      // 这一句很关键，只有notifysuccess后，窗口会被关闭，然后继续后续的操作
      microsoftTeams.authentication.notifySuccess(hashParams["access_token"]);
    } else {
      microsoftTeams.authentication.notifyFailure("授权失败");
    }
  }, [])

  return <Text content="身份认证结束..."></Text>

}

// 这个组件是主界面
function Home() {

  const [authToken, setAuthToken] = useState<string>();
  const [graphToken, setGraphToken] = useState<string>();
  const [userName, setUserName] = useState<string>();
  const [fileContent, setFileContent] = useState<string>();

  return (
    <Flex column fill gap="gap.medium">
      <Button content="获取Auth Token" onClick={() => {
        // 这个按钮用来获取本地的客户端凭据
        microsoftTeams.authentication.getAuthToken({
          successCallback: (token: string) => {
            setAuthToken(token);
          }
        })
      }}></Button>
      <Segment content={authToken} color="red"></Segment>


      <Button content="获取Graph Token" onClick={async () => {
        // 这个按钮用来获取交换得到的graph 令牌
        let serverURL = `api/token?ssoToken=${authToken}&scope=${scope}`;
        let response = await fetch(serverURL);
        if (response) {
          let data = await response.json();
          if (!response.ok && data.error === '要求授权') {
            microsoftTeams.authentication.authenticate({
              url: window.location.origin + "/auth-start",
              width: 600,
              height: 535,
              successCallback: (result) => { setGraphToken(result); },
              failureCallback: (reason) => { console.log(`交换token失败，原因是:${reason}`) }
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
          // 这里拿到的graphToken，可以用来继续访问其他资源
          const client = graph.Client.init({
            authProvider: (done: any) => {
              done(null, graphToken);
            }
          })

          const user = await client.api("/me").get();
          setUserName(user.displayName);

        }
      }}></Button>
      <Segment content={userName} color="green"></Segment>
      <Button content="获取文件内容" onClick={async () => {
        //实现文件读取
        if (graphToken) {
          const client = graph.Client.init({
            authProvider: (done: any) => {
              done(null, graphToken);
            }
          });

          const file = await client.api("/me/drive/root:/demo.txt").get();
          const url = file["@microsoft.graph.downloadUrl"];
          fetch(url).then(value => value.text()).then(text => setFileContent(text));


        }

      }}></Button>
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
