import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneField,
  PropertyPaneLabel,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFx, SPFI } from "@pnp/sp";

export interface IImageOverlayWebPartProps {
  dataSource: string;  
  imageSource: string;
  imageCss: string;
  propertyFilter: string;
  statusProperty: string;
  statuses: string[];
  colors: string[];
  format: string;
  filterValues: string[];
  lefts: string[];
  tops: string[];
  customTextStyles: string[];
}

export default class ImageOverlayWebPart extends BaseClientSideWebPart<IImageOverlayWebPartProps> {

  sp: SPFI;

  public async render(): Promise<void> {

    let itemsString = "";

    try {
      const response = await fetch(this.properties.dataSource, { mode: "cors" });    
      const data = await response.json();
      itemsString = data.filter((item: any) => this.displayItem(item))
        .map((item: any) => this.itemToHtml(item))
        .join("");

      this.domElement.innerHTML = `<div style="display: flex"><img src="${this.properties.imageSource}" style="${this.properties.imageCss}">${itemsString}</div>`;
    }
    catch (ex) {
      console.log(ex);
      this.domElement.innerHTML = `An error has occured while rendering the webpart.`;
    }
  }

  private displayItem(item: any): boolean {
    return this.properties.statuses.indexOf(item[this.properties.statusProperty]) !== -1 && this.properties.filterValues.indexOf(item[this.properties.propertyFilter]) !== -1;
  }

  private itemToHtml(item: any): string {
    const filterIndex = this.properties.filterValues.indexOf(item[this.properties.propertyFilter]);
    const color = this.properties.colors[this.properties.statuses.indexOf(item[this.properties.statusProperty])];
    const left = this.properties.lefts[filterIndex];
    const top = this.properties.tops[filterIndex];
    const style = this.properties.customTextStyles[filterIndex];
    let text = this.properties.format;
    Object.keys(item).forEach((key: string) => {
      text = text.replace(new RegExp('\\{' + key + '\\}', 'g'), item[key]);
    });
    
    return `<div style="position: absolute; left: ${left}; top: ${top};"><pre style="color: ${color}; ${style}">${text}</pre></div>`;
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    this.sp = spfi().using(SPFx(this.context));
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    if (!this.properties.statuses) this.properties.statuses = [];
    if (!this.properties.colors) this.properties.colors = [];

    let statusFields: IPropertyPaneField<any>[] = [];
    statusFields.push(PropertyPaneTextField('statusProperty', { label: "Status property" }));
    for (let i = 0; i <= this.properties.statuses.length && i <= this.properties.colors.length; i++)
    {
      statusFields.push(PropertyPaneLabel("", { text: (i + 1).toString() }));
      statusFields.push(PropertyPaneTextField('statuses[' + i + ']', { label: "Status" }));
      statusFields.push(PropertyPaneTextField('colors[' + i + ']', { label: "Color" }));
    }

    if (!this.properties.filterValues) this.properties.filterValues = [];
    if (!this.properties.lefts) this.properties.lefts = [];
    if (!this.properties.tops) this.properties.tops = [];
    if (!this.properties.customTextStyles) this.properties.customTextStyles = [];

    let itemFields: IPropertyPaneField<any>[] = [];
    itemFields.push(PropertyPaneTextField('propertyFilter', {
      label: "Property filter"
    }));
    itemFields.push(PropertyPaneTextField('format', {
      label: "Format"
    }));
    for (let i = 0; i <= this.properties.filterValues.length && i <= this.properties.lefts.length && i <= this.properties.tops.length; i++)
    {
      itemFields.push(PropertyPaneLabel("", { text: (i + 1).toString() }));
      itemFields.push(PropertyPaneTextField('filterValues[' + i + ']', { label: "Value" }));
      itemFields.push(PropertyPaneTextField('lefts[' + i + ']', { label: "Left" }));
      itemFields.push(PropertyPaneTextField('tops[' + i + ']', { label: "Top" }));
      itemFields.push(PropertyPaneTextField('customTextStyles[' + i + ']', { label: "Custom style" }));
    }


    return {
      pages: [
        {
          header: {
            description: "Test"
          },
          groups: [
            {
              groupName: "Basic",
              groupFields: [
                PropertyPaneTextField('dataSource', {
                  label: "Data source"
                }),
                PropertyPaneTextField('imageSource', {
                  label: "Image source"
                }),
                PropertyPaneTextField('imageCss', {
                  label: "Image style",
                  multiline: true
                })                
              ]
            },
            {
              groupName: "Statuses",
              groupFields: statusFields
            },
            {
              groupName: "Items",
              groupFields: itemFields
            }
          ]
        }
      ]
    };
  }
}
