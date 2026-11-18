import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
export interface IChatExtensionApplicationCustomizerProperties {
    /** Absolute URL of the hub site hosting the config list. Manifest: "hubSiteUrl" */
    hubSiteUrl: string;
    /** SP list name on the hub site. Manifest: "configListTitle" */
    configListTitle?: string;
    /** Title of the fallback list row. Manifest: "defaultTitleMatch" */
    defaultTitleMatch?: string;
}
export default class ChatExtensionApplicationCustomizer extends BaseApplicationCustomizer<IChatExtensionApplicationCustomizerProperties> {
    private _bottomPlaceholder;
    private _container;
    onInit(): Promise<void>;
    private _renderWidget;
    private _onDispose;
}
//# sourceMappingURL=ChatExtensionApplicationCustomizer.d.ts.map