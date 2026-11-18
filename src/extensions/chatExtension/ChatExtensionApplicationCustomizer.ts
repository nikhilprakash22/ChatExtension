import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { ChatWidget } from './components/ChatWidget';

export interface IChatExtensionApplicationCustomizerProperties {
  /** Absolute URL of the hub site hosting the config list. Manifest: "hubSiteUrl" */
  hubSiteUrl: string;
  /** SP list name on the hub site. Manifest: "configListTitle" */
  configListTitle?: string;
  /** Title of the fallback list row. Manifest: "defaultTitleMatch" */
  defaultTitleMatch?: string;
}

export default class ChatExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<IChatExtensionApplicationCustomizerProperties> {

  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _container: HTMLDivElement | undefined;

  @override
  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this._renderWidget);
    this._renderWidget();
    return Promise.resolve();
  }

  private _renderWidget(): void {
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose.bind(this) }
      );
    }

    if (!this._bottomPlaceholder) { return; }

    if (!this._container) {
      this._container = document.createElement('div');
      this._container.id = 'ce-chat-root';
      this._bottomPlaceholder.domElement.appendChild(this._container);
    }

    const hubSiteUrl: string        = this.properties.hubSiteUrl        || 'https://leadecservices.sharepoint.com/sites/VISBSharePointTestSite';
    const listTitle: string         = this.properties.configListTitle   || 'ChatExtensionConfig';
    const defaultTitleMatch: string = this.properties.defaultTitleMatch || 'Default Settings';

    // SiteId column in the SP list stores the site-collection GUID (not a URL).
    // pageContext.site.id is the Guid object — call toString() to get the raw GUID string.
    const siteId: string = this.context.pageContext.site.id
      ? this.context.pageContext.site.id.toString()
      : '';

    // siteUrl is still needed by RbacService for SP REST calls
    const siteUrl: string = this.context.pageContext.web.absoluteUrl || '';

    if (!hubSiteUrl) { return; }

    ReactDOM.render(
      React.createElement(ChatWidget, {
        spHttpClient:       this.context.spHttpClient,
        hubSiteUrl,
        listTitle,
        defaultTitleMatch,
        siteId,
        siteUrl,
      }),
      this._container
    );
  }

  private _onDispose(): void {
    if (this._container) {
      ReactDOM.unmountComponentAtNode(this._container);
    }
  }
}