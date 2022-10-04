import { Log } from '@microsoft/sp-core-library';
import styles from './ImportantCompanyAnnouncementsApplicationCustomizer.module.scss';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} 
from '@microsoft/sp-application-base';
import * as strings from 'ImportantCompanyAnnouncementsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'ImportantCompanyAnnouncementsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IImportantCompanyAnnouncementsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ImportantCompanyAnnouncementsApplicationCustomizer
  extends BaseApplicationCustomizer<IImportantCompanyAnnouncementsApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    return Promise.resolve();
  }
  private _renderPlaceHolders(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top
      );

      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }
      this.context.spHttpClient
        .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Announcements')/items?$select=Title,Description`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'accept': 'application/json;odata.metadata=none'
            }
          })
        .then(response => response.json())
        .then(announcements => {
          const announcementsHtml = announcements.value.map((announcement: { Title: string; Description: string; }) =>
            `<li>${announcement.Title}</li>
             <li>${announcement.Description}`);
          this._topPlaceholder.domElement.innerHTML = `<div class="${styles.app}">
             <ul>${announcementsHtml.join('')}</ul></div>`;
        }).catch(error => console.log(error));
    }
  }
}
