import Axios from "axios";
import { Request } from "express";
import { ConnectorDeclaration, IConnector, PreventIframe } from "express-msteams-host";
import * as debug from "debug";
const JsonDB = require("../../../node_modules/@types/node-json-db");

const log = debug("msteams");

/**
 * The connector data interface
 */
interface IPerksweetConnectorData {
    webhookUrl: string;
    user: string;
    appType: string;
    groupName: string;
    existing: boolean;
}

/**
 * Implementation of the "PerksweetConnectorConnector" Office 365 Connector
 */
@ConnectorDeclaration(
    "/api/connector/connect",
    "/api/connector/ping"
)
@PreventIframe("/perksweetConnector/config.html")
export class PerksweetConnector implements IConnector {
    private connectors: any;

    public constructor() {
        // Instantiate the node-json-db database (connectors.json)
        this.connectors = new JsonDB("connectors", true, false);
    }

    public async Connect(req: any) {
        try {

            const request = {
                email: req.body.email,
                password: req.body.password,
                webhook: req.body.webhookUrl
            };

            return await Axios.post(
                process.env.BACKEND_ENDPOINT ?? "",
                request
            )
                .then(response => {
                    console.log("Response: ", response);
                    console.log(`Response from Connector endpoint is: ${response.status}`);

                    if (response.status === 200 || response.status === 302) {
                        this.connectors.push("/connectors[]", {
                            appType: req.body.appType,
                            existing: true,
                            groupName: req.body.groupName,
                            user: req.body.user,
                            webhookUrl: req.body.webhookUrl
                        });
                    }

                    return response;
                }).catch(error => {
                    if (error.response) {
                        console.log("axios - catch", error.response);

                        console.log(error.response.data);
                        console.log(error.response.status);
                        console.log(error.response.headers);

                        throw new Error("Invalid Credentials");
                    } else {
                        console.log("axios - catch-else", error);

                        throw new Error("Invalid Credentials");
                    }
                });
        } catch (e) {
            console.log("axios - outer-catch", e);
            throw new Error("Invalid Credentials");
        }
    }

    public Ping(req: Request): Array<Promise<void>> {
        // clean up connectors marked to be deleted
        try {
            this.connectors.push("/connectors",
                (this.connectors.getData("/connectors") as IPerksweetConnectorData[])
                    .filter(c => {
                        return c.existing;
                    }));
        } catch (error) {
            if (error.name && error.name === "DataError") {
                // there"s no registered connectors
                return [];
            }
            throw error;
        }

        // send pings to all subscribers
        return (this.connectors.getData("/connectors") as IPerksweetConnectorData[]).map((connector, index) => {
            return new Promise<void>((resolve, reject) => {
                // TODO: implement adaptive cards when supported
                const card = {
                    title: "Sample Connector",
                    text: "This is a sample Office 365 Connector",
                    sections: [
                        {
                            activityTitle: "Ping",
                            activityText: "Sample ping ",
                            activityImage: `https://${process.env.PUBLIC_HOSTNAME}/assets/icon.png`,
                            facts: [
                                {
                                    name: "Generator",
                                    value: "teams"
                                },
                                {
                                    name: "Created by",
                                    value: connector.user
                                }
                            ]
                        }
                    ],
                    potentialAction: [{
                        "@context": "http://schema.org",
                        "@type": "ViewAction",
                        name: "Yo Teams",
                        target: ["http://aka.ms/yoteams"]
                    }]
                };

                log(`Sending card to ${connector.webhookUrl}`);

                Axios.post(
                    decodeURI(connector.webhookUrl),
                    JSON.stringify(card)
                ).then(response => {
                    log(`Response from Connector endpoint is: ${response.status}`);
                    resolve();
                }).catch(err => {
                    if (err.response && err.response.status === 410) {
                        this.connectors.push(`/connectors[${index}]/existing`, false);
                        log(`Response from Connector endpoint is: ${err.response.status}, add Connector for removal`);
                        resolve();
                    } else {
                        log(`Error from Connector endpoint is: ${err}`);
                        reject(err);
                    }
                });
            });
        });
    }
}
