import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
interface IChatWidgetProps {
    spHttpClient: SPHttpClient;
    hubSiteUrl: string;
    listTitle: string;
    defaultTitleMatch: string;
    siteId: string;
    siteUrl: string;
    userRole?: string;
}
export declare const ChatWidget: React.FC<IChatWidgetProps>;
export {};
//# sourceMappingURL=ChatWidget.d.ts.map