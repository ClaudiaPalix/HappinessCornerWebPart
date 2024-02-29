import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HappinessCornerWebPart.module.scss';
import * as strings from 'HappinessCornerWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IHappinessCornerWebPartProps {
  description: string;
  selectedList: string; 
  Title: string;
  URL: string;
  Image: {
    Url: string;
  };
}

export default class HappinessCornerWebPart extends BaseClientSideWebPart<IHappinessCornerWebPartProps> {
  private availableLists: IPropertyPaneDropdownOption[] = [];


  public render(): void {
    const decodedDescription = decodeURIComponent(this.properties.description); // Decode the description (like incase there is blank space, or special characters, etc)
    console.log(decodedDescription);
    this.domElement.innerHTML = `
      <div class="${styles['group-information']}">
        <div class="${styles['parent-div']} ">
        <div class="${styles.topSection}">
        <h2>${(decodedDescription)}</h2> <!-- Use the decoded description -->
        <!--<a href="https://www.google.com" target="_blank">See all</a>-->
        </div>
        <div id="buttonsContainer">
        </div>
        </div>
      </div>`;

    this._renderButtons();
 
  }

  private _renderButtons(): void {
    const buttonsContainer: HTMLElement | null = this.domElement.querySelector('#buttonsContainer');
    buttonsContainer?.classList.add(styles.buttonsContainer);
    const apiUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.selectedList}')/items`;
    
    const siteUrl = this.context.pageContext.web.absoluteUrl;

    fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    })
    .then(response => response.json())
    .then(data => {
      console.log("API Response: ", data); // Log the entire response for inspection

      if (data.value && data.value.length > 0) {
        data.value.forEach((item: IHappinessCornerWebPartProps) => {
          const colDiv: HTMLDivElement = document.createElement('div');
          const innerDiv: HTMLDivElement = document.createElement('div');
          const buttonDiv: HTMLDivElement = document.createElement('div');
          colDiv.classList.add(styles['colDiv'] , "colDiv");
          innerDiv.classList.add(styles['content-section']);
          buttonDiv.classList.add(styles['content-box']); // Apply the 'content-box' class styles from YourStyles.module.scss
          buttonDiv.onclick = (event) => {
            if(item.URL.includes(siteUrl)){
              event.preventDefault(); // Prevent the default behavior of the click event
              window.location.href = item.URL; // Navigate to the 'URL' from the API response in the same tab
            }else{
              window.open(item.URL, '_blank'); // Open the 'Url' from the API response in a new tab
            }
          };

          const imgContainer: HTMLDivElement = document.createElement('div');
          imgContainer.classList.add(styles['content-box-img-container']); // Apply the image container styles
          const img: HTMLImageElement = document.createElement('img');
          img.src = item.Image.Url; // Use the imported image URL
          imgContainer.appendChild(img); // Append the image to the container
          buttonDiv.appendChild(imgContainer); // Append the image container to the button

          const textDiv: HTMLDivElement = document.createElement('div');
          textDiv.classList.add(styles['content-box-text-container']);
          const titleSpan: HTMLDivElement = document.createElement('div');
          titleSpan.textContent = item.Title; // Use the 'Title' from the API response
          textDiv.appendChild(titleSpan); // Append the title to the button
          
          colDiv.appendChild(innerDiv);
          innerDiv.appendChild(buttonDiv);
          innerDiv.appendChild(textDiv);
          buttonsContainer!.appendChild(colDiv); 
        });
      } else {
        const noDataMessage: HTMLDivElement = document.createElement('div');
        noDataMessage.textContent = 'No applications available for the user.';
        buttonsContainer!.appendChild(noDataMessage);// Non-null assertion operator
      }

      
    
     // Use 'this.domElement' to find the dynamically created elements
     const HappinessCornerParent = this.domElement.querySelector(`.${styles['parent-div']}`);
     const colDiv = this.domElement.querySelectorAll(".colDiv");
     const containerDiv = this.domElement.querySelectorAll("#buttonsContainer");
 
     const handleResponsiveCheck = () => {
       if (HappinessCornerParent) {
         const windowWidth = HappinessCornerParent.getBoundingClientRect().width;
 
         if (windowWidth < 400) {
           colDiv.forEach(element => {
             element.classList.add(styles.ColSmallerWidthSm);
             element.classList.remove(styles.ColSmallerWidthMd);
           });
           containerDiv.forEach(element => {
            element.classList.add(styles.ContainerSmallerWidthSm);
            element.classList.remove(styles.ContainerSmallerWidthMd);
          });
         } else if (windowWidth < 630) {
           colDiv.forEach(element => {
             element.classList.add(styles.ColSmallerWidthMd);
             element.classList.remove(styles.ColSmallerWidthSm);
           });
           containerDiv.forEach(element => {
            element.classList.add(styles.ContainerSmallerWidthMd);
            element.classList.remove(styles.ContainerSmallerWidthSm);
          });
         } else {
           colDiv.forEach(element => {
             element.classList.remove(styles.ColSmallerWidthSm, styles.ColSmallerWidthMd);
           });
         }
       }
     };
 
     handleResponsiveCheck();
     window.addEventListener('resize', handleResponsiveCheck);
    })


    .catch(error => {
      console.error("Error fetching user data: ", error);
    });


  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._loadLists();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'selectedList') {
      this.setListTitle(newValue);
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private _loadLists(): void {
    const listsUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`;
    //SPHttpClient is a class provided by Microsoft that allows developers to perform HTTP requests to SharePoint REST APIs or other endpoints within SharePoint or the host environment. It is used for making asynchronous network requests to SharePoint or other APIs in SharePoint Framework web parts, extensions, or other components.
    this.context.spHttpClient.get(listsUrl, SPHttpClient.configurations.v1)
    //SPHttpClientResponse is the response object returned after making a request using SPHttpClient. It contains information about the response, such as status code, headers, and the response body.
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value: any[] }) => {
        this.availableLists = data.value.map((list) => {
          return { key: list.Title, text: list.Title };
        });
        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        console.error('Error fetching lists:', error);
      });
  }

  private setListTitle(selectedList: string): void {
    this.properties.selectedList = selectedList;

    this.context.propertyPane.refresh();
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
                  label: "Title For The Application"
                }),
                PropertyPaneDropdown('selectedList', {
                  label: 'Select A List',
                  options: this.availableLists,
                }),
              ],
            },
          ],
        }
      ]
    };
  }
}

