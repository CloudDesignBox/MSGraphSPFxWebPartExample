import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient } from '@microsoft/sp-client-preview';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import styles from './MyOneDriveFilesWebPart.module.scss';
import * as strings from 'MyOneDriveFilesWebPartStrings';

export interface IMyOneDriveFilesWebPartProps {
  description: string;
}

export default class MyOneDriveFilesWebPart extends BaseClientSideWebPart<IMyOneDriveFilesWebPartProps> {

  public render(): void {
    let htmlcode:string = "";
    const client: MSGraphClient = this.context.serviceScope.consume(MSGraphClient.serviceKey);
    
    client
      .api('me/drive/recent')
      .get((error, files: MicrosoftGraph.DriveItem, rawResponse: any) => {
        // handle the response
        for (var _i = 0; _i < rawResponse.body.value.length; _i++) {
          htmlcode += `<a href="${rawResponse.body.value[_i].webUrl}">${rawResponse.body.value[_i].name}</a></br>`;

        }
      this.domElement.innerHTML = `
      <div class="${ styles.myOneDriveFiles }">
        ${htmlcode}
      </div>`;
    });
    
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
