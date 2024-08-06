import * as React from "react";
import { Provider, Flex, Input } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import Axios from "axios";

/**
 * Implementation of the connectorConnector Connector connect page
 */
export const PerksweetConnectorConfig = () => {
    const [{ theme, context }] = useTeams();
    const [inputs, setInputs] = useState({
        email: "",
        password: ""
    });
    const handleChange = (event) => {
        const name = event.target.name;
        const value = event.target.value;
        setInputs(values => ({ ...values, [name]: value }));
    };

    useEffect(() => {
        console.log(inputs);
        if (context) {
            microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
                // INFO: Should really be of type microsoftTeams.settings.Settings, but configName does not exist in the Teams JS SDK
                const settings: any = {
                    email: inputs.email,
                    password: inputs.password,
                    contentUrl: "https://teams.perksweet.com/connectorConnector/config.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}"
                };

                microsoftTeams.settings.setSettings(settings);

                microsoftTeams.settings.getSettings((setting: any) => {
                    try {
                        const request = {
                            email: settings.email,
                            password: settings.password,
                            webhook: setting.webhookUrl
                        };

                        Axios.post(
                            "https://perksweet.com/api/configure-teams-webhook",
                            request
                        )
                            .then(response => {
                                console.log("Response: ", response);
                                console.log(`Response from Connector endpoint is: ${response.status}`);

                                if (response.status === 200 || response.status === 302) {
                                    saveEvent.notifySuccess();
                                } else {
                                    saveEvent.notifyFailure("Invalid Credentials");
                                }

                            }).catch(error => {
                                if (error.response) {
                                    console.log("axios - catch", error.response);

                                    saveEvent.notifyFailure("Invalid Credentials");
                                } else {
                                    console.log("axios - catch-else", error);

                                    saveEvent.notifyFailure("Invalid Credentials");
                                }
                            });

                        // fetch("/api/connector/connect", {
                        //     method: "POST",
                        //     headers: [
                        //         ["Content-Type", "application/json"]
                        //     ],
                        //     body: JSON.stringify({
                        //         webhookUrl: setting.webhookUrl,
                        //         user: setting.userObjectId,
                        //         appType: setting.appType,
                        //         groupName: context.groupId,
                        //         email: settings.email,
                        //         password: settings.password
                        //     })
                        // }).then(response => {
                        //     console.log("response in fetch-then : ", response);
                        //
                        //     if (response.statusText === 'Internal Server Error') {
                        //         saveEvent.notifyFailure('Invalid Credentials');
                        //         return;
                        //     }
                        //
                        //     if ((response.status === 200 || response.status === 302)) {
                        //         saveEvent.notifySuccess();
                        //     } else {
                        //         saveEvent.notifyFailure(response.statusText);
                        //     }
                        // }).catch(e => {
                        //     console.log("response in fetch-catch : ", e);
                        //     saveEvent.notifyFailure(e);
                        // });
                    } catch (e) {
                        console.log("response in fetch-outer-catch", e);
                        saveEvent.notifyFailure(e);
                    }
                });
            });

            microsoftTeams.settings.registerOnRemoveHandler((RemoveEvent: microsoftTeams.settings.RemoveEvent) => {
                microsoftTeams.settings.getSettings((setting: any) => {
                    console.log("setting : ", setting, RemoveEvent);

                    const request = {
                        email: setting.email,
                        password: setting.password,
                        webhook: setting.webhookUrl
                    };

                    Axios.post(
                        process.env.BACKEND_REMOVE_ENDPOINT ?? "",
                        request
                    )
                        .then(response => {
                            if (response.status === 200 || response.status === 302) {
                                RemoveEvent.notifySuccess();
                            } else {
                                RemoveEvent.notifyFailure(response.statusText);
                            }
                        }).catch(e => {
                            console.log(e.response.data);
                            console.log(e.response.status);
                            console.log(e.response.headers);

                            RemoveEvent.notifyFailure(e);

                            throw new Error(e.response.status);
                        });
                });
            });
        }
    }, [inputs.email, inputs.password, context]);

    useEffect(() => {
        if (context) {
            let validityState = false;
            const regex = /^(([^<>()[\].,;:\s@"]+(\.[^<>()[\].,;:\s@"]+)*)|(".+"))@(([^<>()[\].,;:\s@"]+\.)+[^<>()[\].,;:\s@"]{2,})$/i;

            if ((!inputs.email || regex.test(inputs.email)) && inputs.password) {
                validityState = true;
            }

            microsoftTeams.settings.setValidityState(validityState);
        }
    }, [inputs.email, inputs.password, context]);

    return (
        <Provider theme={theme}>
            <Flex fill={true} column={true}>
                <Flex.Item>
                    <div>
                        <p style={{ fontWeight: "bold" }}>
                            Important Notes:
                        </p>

                        <ul style={{ marginBlockStart: "-8px" }}>
                            <li>
                                No need to generate incoming webhook manually, just configure connector by using PerkSweet credentials.
                            </li>
                            <li>
                                If you dont have account with PerkSweet, then please <a href="https://perksweet.com/register/company" target="_blank" rel="noopener noreferrer">Sign Up</a> here. after signup, use that credentials to configure connector. so perksweet can post all updates of your organizations directly to your MS-Teams channel.
                            </li>
                            <li>
                                This app will post update, whenever someone send kudos to another person in your organization, birthday updates & anniversary updates.
                            </li>
                        </ul>

                        <h2>
                            Configure your Connector using PerkSweet Credentials
                        </h2>

                        <Input type="email"
                            name="email"
                            value={inputs.email}
                            required={true}
                            placeholder="Email"
                            onChange={handleChange}/>

                        <br/><br/>

                        <Input type="password"
                            name="password"
                            value={inputs.password}
                            required={true}
                            placeholder="Password"
                            onChange={handleChange}/>
                    </div>
                </Flex.Item>

                <Flex.Item>
                    <div>
                        <p>
                            Don&apos;t have an account ? <a href="https://perksweet.com/register/company" target="_blank" rel="noopener noreferrer">Sign Up</a>
                        </p>
                    </div>
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
