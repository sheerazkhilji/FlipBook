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
debugger;
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
    const prevButton = this.domElement.querySelector('#prevPage') as HTMLButtonElement;
    const nextButton = this.domElement.querySelector('#nextPage') as HTMLButtonElement;
    const flipbookContainer = this.domElement.querySelector('#flipbookContainer') as HTMLElement; // Assuming there's a container with this ID

    if (flipbookElement && prevButton && nextButton) {
        // Center the flipbook on initialization
      //  this.centerFlipbook(flipbookElement);

        // Initialize the flipbook with jQuery
        (jQuery as any)(flipbookElement).turn({
            width: 600,         // Increased width for better viewing
            height: 450,        // Increased height for better viewing
            autoCenter: true,  // Set to false to manually handle centering
            display: 'double',   // Double-page view for a realistic book effect
            duration: 1000,     // Slows down the page-turning animation for visual appeal
            elevation: 50,      // Adds 3D depth to the flip effect
            // Add navigation if available
            next: () => {
                (jQuery as any)(flipbookElement).turn('next');
                this.checkLastPageAndCenter(flipbookElement, flipbookContainer);
            },
            previous: () => {
                (jQuery as any)(flipbookElement).turn('previous');
            },
            when: {
                // Callback when the flipbook is closed
                closed: () => {
                    // Center the flipbook when it is closed
               //     this.centerFlipbook(flipbookElement);
                }
            }
        });

        // Go to the previous page
        prevButton.addEventListener('click', () => {
            (jQuery as any)(flipbookElement).turn('previous');
            this.checkLastPageAndCenter(flipbookElement, flipbookContainer);
        });

        // Go to the next page
        nextButton.addEventListener('click', () => {
            (jQuery as any)(flipbookElement).turn('next');
            this.checkLastPageAndCenter(flipbookElement, flipbookContainer);
        });


    } else {
        console.error('Flipbook element or navigation buttons not found');
    }
}

// Method to check if the current page is the last page and center the flipbook
private checkLastPageAndCenter(flipbookElement: HTMLElement, flipbookContainer: HTMLElement): void {
    const totalPages = (jQuery as any)(flipbookElement).turn('pages');
    const currentPage = (jQuery as any)(flipbookElement).turn('page');
debugger;
    if (currentPage === totalPages) {
        // Center the flipbook when on the last page
     //   this.centerFlipbook(flipbookElement);
        // Set flipbookContainer to left: 200px
        flipbookElement.style.left = '200px';
    } else {
        // Reset to default left positioning if not on the last page
        flipbookElement.style.left = ''; // or set to a specific default value if needed
    }
}

// Method to center the flipbook
// private centerFlipbook(flipbookElement: HTMLElement): void {
//     const container = flipbookElement.parentElement; // Get the container of the flipbook
//     if (container) {
//         const containerWidth = container.clientWidth;
//         const containerHeight = container.clientHeight;
//         const flipbookWidth = flipbookElement.offsetWidth;
//         const flipbookHeight = flipbookElement.offsetHeight;

//         // Calculate the new position for centering
//         const left = (containerWidth - flipbookWidth) / 2;
//         const top = (containerHeight - flipbookHeight) / 2;

//         // Set the position of the flipbook
//         flipbookElement.style.left = `${left}px`;
//         flipbookElement.style.top = `${top}px`;
//     }
// }


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
