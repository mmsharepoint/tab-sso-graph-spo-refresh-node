import React from "react";
import { app, pages } from "@microsoft/teams-js";
import config from "./sample/lib/config";

export const Config: React.FC<{}> = () =>  {
  //React.useEffect(() => {
    console.log(`${process.env.REACT_APP_TAB_ENDPOINT}/EnsureSiteUser`);
    app.initialize()
      .then(() => {
        pages.config.setValidityState(true);
        pages.config.registerOnSaveHandler((saveEvent) => {
          const configPromise = pages.config.setConfig({
            websiteUrl: `${process.env.REACT_APP_TAB_ENDPOINT}/EnsureSiteUser`,
            contentUrl: `${process.env.REACT_APP_TAB_ENDPOINT}/EnsureSiteUser`,
            entityId: "grayIconTab",
            suggestedDisplayName: "Ensure Site User"
          });
          configPromise.
            then((result) => {saveEvent.notifySuccess()}).
            catch((error) => {saveEvent.notifyFailure("failure message")});
        });
      })
      .catch((ex) => {
        console.log(ex);
      });    
  //}, []);
  return (
    <div className="welcome page">
      <h1>Configure Teams Tab</h1>
    </div>
  );
}