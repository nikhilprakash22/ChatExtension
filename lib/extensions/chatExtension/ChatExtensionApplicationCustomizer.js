var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer, PlaceholderName, } from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { ChatWidget } from './components/ChatWidget';
var ChatExtensionApplicationCustomizer = /** @class */ (function (_super) {
    __extends(ChatExtensionApplicationCustomizer, _super);
    function ChatExtensionApplicationCustomizer() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ChatExtensionApplicationCustomizer.prototype.onInit = function () {
        this.context.placeholderProvider.changedEvent.add(this, this._renderWidget);
        this._renderWidget();
        return Promise.resolve();
    };
    ChatExtensionApplicationCustomizer.prototype._renderWidget = function () {
        if (!this._bottomPlaceholder) {
            this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom, { onDispose: this._onDispose.bind(this) });
        }
        if (!this._bottomPlaceholder) {
            return;
        }
        if (!this._container) {
            this._container = document.createElement('div');
            this._container.id = 'ce-chat-root';
            this._bottomPlaceholder.domElement.appendChild(this._container);
        }
        var hubSiteUrl = this.properties.hubSiteUrl || 'https://leadecservices.sharepoint.com/sites/VISBSharePointTestSite';
        var listTitle = this.properties.configListTitle || 'ChatExtensionConfig';
        var defaultTitleMatch = this.properties.defaultTitleMatch || 'Default Settings';
        // SiteId column in the SP list stores the site-collection GUID (not a URL).
        // pageContext.site.id is the Guid object — call toString() to get the raw GUID string.
        var siteId = this.context.pageContext.site.id
            ? this.context.pageContext.site.id.toString()
            : '';
        // siteUrl is still needed by RbacService for SP REST calls
        var siteUrl = this.context.pageContext.web.absoluteUrl || '';
        if (!hubSiteUrl) {
            return;
        }
        ReactDOM.render(React.createElement(ChatWidget, {
            spHttpClient: this.context.spHttpClient,
            hubSiteUrl: hubSiteUrl,
            listTitle: listTitle,
            defaultTitleMatch: defaultTitleMatch,
            siteId: siteId,
            siteUrl: siteUrl,
        }), this._container);
    };
    ChatExtensionApplicationCustomizer.prototype._onDispose = function () {
        if (this._container) {
            ReactDOM.unmountComponentAtNode(this._container);
        }
    };
    __decorate([
        override
    ], ChatExtensionApplicationCustomizer.prototype, "onInit", null);
    return ChatExtensionApplicationCustomizer;
}(BaseApplicationCustomizer));
export default ChatExtensionApplicationCustomizer;
//# sourceMappingURL=ChatExtensionApplicationCustomizer.js.map