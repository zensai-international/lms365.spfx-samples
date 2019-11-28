import { Version } from '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './Lms365StatisticsWebPart.module.scss';
import * as strings from 'Lms365StatisticsWebPartStrings';
import { AadHttpClient } from '@microsoft/sp-http';

export interface ILms365StatisticsWebPartProps {
    description: string;
}

export default class Lms365StatisticsWebPart extends BaseClientSideWebPart<ILms365StatisticsWebPartProps> {

    private LMS365APIResourceId = '751987ae-7307-451f-a4b1-e4dc2dfdd507';

    public async render() {
        const lmsAPIClient = await this.context.aadHttpClientFactory.getClient(this.LMS365APIResourceId);
        const response = await lmsAPIClient.get('https://api.365.systems/odata/v2/Courses?$count=true&$top=0', AadHttpClient.configurations.v1);
        const responseJson = await response.json();
        //please use region based uris
        const enrollmentsResponse = await lmsAPIClient.get('https://api.365.systems/odata/v2/Enrollments?$count=true&$top=0', AadHttpClient.configurations.v1);
        const enrollmentsResponseJson = await enrollmentsResponse.json();
        
    this.domElement.innerHTML = `
    <div class="${ styles.lms365Statistics }">
      <div class="${ styles.container }">
      <div class="${ styles.row }">
          <div class="${ styles.column }">
            <span class="${ styles.title }">LMS365 Small Stats</span>         
          </div>
        </div>
        <div class="${ styles.row }">
          <div class="${ styles.column }">
            <span class="${ styles.title }">Courses I can see</span>
            <p class="${ styles.subTitle }">${responseJson['@odata.count']}</p>           
         
          </div>
        </div>
        <div class="${ styles.row }">
          <div class="${ styles.column }">
            <span class="${ styles.title }">Enrollments I can see</span>
            <p class="${ styles.subTitle }">${enrollmentsResponseJson['@odata.count']}</p>
          </div>
        </div>
      </div>
    </div>`;
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
