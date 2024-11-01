declare interface IFbWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}


declare module 'FbWebPartStrings' {
  const strings: IFbWebPartStrings;
  export = strings;
}
declare module '*.module.scss' {
  const styles: { [className: string]: string };
  export default styles;
}

declare module 'jquery' {
    interface JQuery<TElement = HTMLElement> {
        turn(options?: any): JQuery<TElement>;
    }
}