declare interface IFastFiltersWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'FastFiltersWebPartStrings' {
  const strings: IFastFiltersWebPartStrings;
  export = strings;
}
