import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

// import styles from './NewsAlertWebPart.module.scss';
import * as strings from 'NewsAlertWebPartStrings';

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Description: string;
  AlertType: any;
  Active: boolean;
  Position: number;
}

export interface INewsAlertWebPartProps {
  description: string;
  listName: string;
  titleFontSize: string;
}

const sortByColumn = <T, K extends keyof T>(arr: T[], columnName: K) => {
  return [...arr].sort((a, b) => {
    if (a[columnName] < b[columnName]) return -1;
    if (a[columnName] > b[columnName]) return 1;
    return 0;
  });
};

export default class NewsAlertWebPart extends BaseClientSideWebPart<INewsAlertWebPartProps> {

  constructor() {
    super();
    this._getListData = this._getListData.bind(this);
    this._renderList = this._renderList.bind(this);
    this._renderListAsync = this._renderListAsync.bind(this);
  }

  private _getListData(): Promise<ISPLists> {
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/Lists/GetByTitle('${this.properties.listName}')/Items`;
    console.log('Fetching list data from:', requestUrl);

    return this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        console.log('Response received:', response);
        return response.json();
      })
      .then((jsonData: ISPLists) => {
        console.log('JSON data received:', jsonData);
        return jsonData;
      })
      .catch((error) => {
        console.error('Error fetching data:', error);
        throw error;
      });
  }

  private _renderList(items: ISPList[]): void {
    console.log('Items before filtering:', items);
    const activeItems = items.filter(item => item.Active === true);
    console.log('Active items:', activeItems);
    const sortedItems = sortByColumn(activeItems, 'Position');
    console.log('Sorted items:', sortedItems);

    let html: string = '<div class="card_parent">';
    sortedItems.forEach((itm: ISPList) => {
      console.log('Processing item before html:', itm);
      html += `
        <div class="image_item">
          <p>${itm.Title}</p>
          <p>${itm.Description}</p>
          <p>${itm.AlertType}</p>
          <p>${itm.Active}</p>
        </div>`;
    });
    html += `</div>`;

    const container = document.querySelector('#spListContainer');
    if (container != null) {
      container.innerHTML = html;
    } else {
      console.error('Container not found!');
    }
  }

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        console.log('Data fetched successfully:', response.value);
        this._renderList(response.value);
      })
      .catch((error) => {
        console.error('Error in _renderListAsync:', error);
      });
  }

  public render(): void {
    this.domElement.innerHTML = `
      <section>
        <div id="spListContainer"></div>
      </section>`;

    this._renderListAsync();
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      console.log('Environment message:', message);
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
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
                PropertyPaneTextField('listName', {
                  label: 'List Name'
                }),
                PropertyPaneTextField('titleFontSize', {
                  label: 'title Font Size'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
