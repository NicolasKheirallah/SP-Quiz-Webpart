declare interface IQuizWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  ListNameFieldLabel: string;
  ResultsListNameFieldLabel: string;
  QuestionsPerPageFieldLabel: string;
  AppLocalEnvironmentOffice: string;
  AppOfficeEnvironment: string;
  AppLocalEnvironmentOutlook: string;
  AppOutlookEnvironment: string;
  AppLocalEnvironmentTeams: string;
  AppTeamsTabEnvironment: string;
  UnknownEnvironment: string;
  AppLocalEnvironmentSharePoint: string;
  AppSharePointEnvironment: string;
  DescriptionFieldLabel: string;
}

declare module 'QuizWebPartStrings' {
  const strings: IQuizWebPartStrings;
  export = strings;
}
