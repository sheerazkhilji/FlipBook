import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'FbWebPartStrings';
import axios from 'axios';
import * as jQuery from 'jquery'; // Namespace import

import 'turn.js'; // Ensure Turn.js is imported
import styles from './FbWebPart.module.scss';

export interface IFbWebPartProps {
  description: string;
}

export default class FbWebPart extends BaseClientSideWebPart<IFbWebPartProps> {
  private images: string[] = []; // Store the image URLs

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.flipbookContainer}">
        <h3>Image Flipbook</h3>
        <div id="flipbook" class="${styles.flipbook}"></div>
      </div>
    `;
    this.loadImagesFromLibrary();
  }

  private loadImagesFromLibrary(): void {
    const libraryName = 'TestFolder'; // Update to your document library name if needed
    const query = `/_api/web/lists/getbytitle('${libraryName}')/items?$select=FileRef,FileLeafRef&$filter=substringof('.jpg',FileLeafRef) or substringof('.png',FileLeafRef) or substringof('.gif',FileLeafRef)`;

    axios.get(this.context.pageContext.web.absoluteUrl + query, {
      headers: {
        Accept: 'application/json;odata=verbose'
      }
    })
    .then(response => {
      const images = response.data.d.results;
      this.images = images.map((img: { FileRef: string; }) => {
        const relativeUrl = img.FileRef.replace(/\/TestSite(\/)?/, '/'); // Remove duplicate "TestSite" if it exists
        return this.context.pageContext.web.absoluteUrl + relativeUrl;
      });
      this.displayImages();
    })
    .catch(error => {
      console.error('Error fetching images:', error);
    });
  }

  private displayImages(): void {
    const flipbookElement = this.domElement.querySelector('#flipbook') as HTMLElement;

    if (flipbookElement) {
      // Clear previous content
      flipbookElement.innerHTML = '';

      this.images.forEach(imageUrl => {
        // Wrap each image in a div with the class 'page'
        flipbookElement.innerHTML += `
          <div class="${styles.page}">
              <img src="${imageUrl}" alt="Image" />
          </div>
        `;
      });

         // Insert navigation buttons after the flipbook
         flipbookElement.insertAdjacentHTML('afterend', `
          <div class="${styles.navigation}">
              <button id="prevPage">Previous</button>
              <button id="nextPage">Next</button>
          </div>
      `);

      
      this.initializeFlipbook();
    }
  }

  private initializeFlipbook(): void {
    const flipbookElement = this.domElement.querySelector('#flipbook') as HTMLElement;

    if (flipbookElement) {
        // Initialize the flipbook with jQuery
        (jQuery as any)(flipbookElement).turn({
            width: 400,
            height: 300,
            autoCenter: true,
            // Add navigation if available
            next: () => (jQuery as any)(flipbookElement).turn('next'),
            previous: () => (jQuery as any)(flipbookElement).turn('previous')
        });

        // Set up navigation buttons
        const prevButton = this.domElement.querySelector('#prevPage') as HTMLButtonElement;
        const nextButton = this.domElement.querySelector('#nextPage') as HTMLButtonElement;

        // Go to the previous page
        prevButton.addEventListener('click', () => {
            (jQuery as any)(flipbookElement).turn('previous');
        });

        // Go to the next page
        nextButton.addEventListener('click', () => {
            (jQuery as any)(flipbookElement).turn('next');
        });
    } else {
        console.error('Flipbook element not found');
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
