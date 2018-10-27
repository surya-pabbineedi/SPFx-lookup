declare interface ILookupWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  DescriptionFieldLabel: string;
  ListFieldLabel: string;
  FieldsLabel: string;
  ShowUnsupportedFieldsLabel: string;
  LocalWorkbenchUnsupported: string;
  MissingListConfiguration: string;
  ConfigureWebpartButtonText: string;
  ErrorOnLoadingLists: string;
}

declare module 'LookupWebPartStrings' {
  const strings: ILookupWebPartStrings;
  export = strings;
}
