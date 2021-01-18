import * as React from "react";
import { Provider, Flex, Header, Input } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * Implementation of Agenda configuration page
 */
export const AgendaTabConfig = () => {

    const [{ inTeams, theme, context }] = useTeams({});
    const [customSetting, setCustomSetting] = useState<string>("");

    useEffect(() => {
        if (context) {

            setCustomSetting(context.entityId);

            microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
                const host = "https://" + window.location.host;
                microsoftTeams.settings.setSettings({
                    contentUrl: host + "/agendaTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}",
                    websiteUrl: host + "/agendaTab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}",
                    suggestedDisplayName: "Agenda",
                    removeUrl: host + "/agendaTab/remove.html?theme={theme}",
                    entityId: customSetting
                });
                saveEvent.notifySuccess();
            });

            microsoftTeams.settings.setValidityState(true);
            microsoftTeams.appInitialization.notifySuccess();
        }
    }, [context]);

    return (
        <Provider theme={theme}>
            <Flex fill={true}>
                <Flex.Item>
                    <div>
                        <Header content="Configure your tab" />
                        <Input
                            placeholder="Enter a value here"
                            fluid
                            clearable
                            value={customSetting}
                            onChange={(e, data) => {
                                if (data) {
                                    setCustomSetting(data.value);
                                }
                            }}
                            required />
                    </div>
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
