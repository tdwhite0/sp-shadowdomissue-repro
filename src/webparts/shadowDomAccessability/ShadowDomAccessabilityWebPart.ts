import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './ShadowDomAccessabilityWebPart.module.scss';
import * as strings from 'ShadowDomAccessabilityWebPartStrings';

export interface IShadowDomAccessabilityWebPartProps {
  description: string;
}

export default class ShadowDomAccessabilityWebPart extends BaseClientSideWebPart<IShadowDomAccessabilityWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    // create a shadowRoot for the component
    const shadowRoot = this.domElement.attachShadow({mode: "open" });

    shadowRoot.innerHTML = `
    <section class="${styles.shadowDomAccessability} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Tab Key Accessibility does not work in SharePoint for elements inside a Shadow DOM component.</h2>
        <div>${this._environmentMessage}</div>
      </div>
      <div>
  
      <div>Try to use the tab key on your keyboard to focus on elements inside this WebPart. You will find that none of them can be focused.</div>

      <div>
        <h4>Can't press tab between these input boxes:</h4>
          <input style="width: 100%" type="text" placeholder="Click inside here and press tab" />
          <input style="width: 100%" type="text" placeholder="You won't be able to get here" />
      </div>

      <div>
      <h4>Can't press tab between these buttons:</h4>
        <button>Try to tab through us</button>
        <button>Try to tab through us</button>
    </div>
    
    <div>
    <h4>Setting tabindex="0" doesn't matter:</h4>
      <input tabindex="0" style="width: 100%" type="text" placeholder="Click inside here and press tab" />
      <input tabindex="0" style="width: 100%" type="text" placeholder="You won't be able to get here" />
  </div>

  <br />
  <div>
    Root cause: SharePoint's accessibility manager code intercepts every tab and tries to determine where the next focusable element is. This code does not take into account anything inside a Shadow Root because document.querySelector by default does not include elements in Shadow Roots. 
  </div>

    </section>`;
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
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

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

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
