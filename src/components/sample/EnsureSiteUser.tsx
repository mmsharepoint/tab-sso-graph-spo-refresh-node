import React from "react";
import { app, appInitialization, authentication } from "@microsoft/teams-js";
import { BearerTokenAuthProvider, createApiClient, TeamsUserCredential } from "@microsoft/teamsfx";
import { Button, Field, Input, InputOnChangeData, Textarea } from "@fluentui/react-components";
import Axios from "axios";
import config from "./lib/config";

export const EnsureSiteUser: React.FC<{}> = () =>  {
  const [siteUrl, setSiteUrl] = React.useState<string>("");
  const [groupId, setGroupId] = React.useState<string>("");
  const [sitePath, setSitePath] = React.useState<string>("");
  const [userLogin, setUserLogin] = React.useState<string>("");
  const [token, setToken] = React.useState<string>("");

  const onUserLoginChange =  React.useCallback((ev: React.ChangeEvent<HTMLInputElement>, newValue: InputOnChangeData) => {
    setUserLogin(newValue.value);
  }, [userLogin]);

  const onEnsureUser = React.useCallback(() => {
    const requestBody = {
      groupId: groupId,
      sitePath: sitePath
    };

    Axios.post(`${config.apiEndpoint}/api/getSPOUser`, requestBody, {
      responseType: "json",
      headers: {
        Authorization: `Bearer ${token}`
      }
    }).then(result => {
      console.log(result);
    })
    .catch((error) => {
      console.log(error);
    });
  }, [token]);

  React.useEffect(() => {
    app.initialize()
    .then(() => {
      app.getContext()
      .then((context) => {
        if (context.sharePointSite !== undefined) {
          setSiteUrl(context.sharePointSite?.teamSiteUrl!);
          setGroupId(context.team?.groupId!);
          setSitePath(context.sharePointSite?.teamSitePath!);
          setUserLogin(context.user?.userPrincipalName!);
        }
        authentication.getAuthToken()
          .then((token) => {
            setToken(token);          
          });
      });
    });  
  }, []);

  return (
    <div className="welcome page">
      <div>
        <Field label='Site Url' >
          <Input disabled  value={siteUrl} />
        </Field>
      </div>
      <div>
        <Field label='Group / Team ID' >
          <Input disabled value={groupId} />
        </Field>
      </div>
      <div>
        <Field label='Site Path' >
          <Input disabled value={sitePath} />
        </Field>
      </div>
      <div>
        <Field label='Site Url' >
          <Input value={userLogin} onChange={onUserLoginChange} />
        </Field>
      </div>
      <div>
        <Field label='Token' >
          <Textarea value={token} disabled />
        </Field>
      </div>
      <div>
        <Button appearance="primary" onClick={onEnsureUser}>Click me</Button>
      </div>
    </div>
  );
}