declare interface IDashboardWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DashboardTypeFieldLabel: string;
  WebUrlFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
}

declare module 'DashboardWebPartStrings' {
  const strings: IDashboardWebPartStrings;
  export = strings;
}
