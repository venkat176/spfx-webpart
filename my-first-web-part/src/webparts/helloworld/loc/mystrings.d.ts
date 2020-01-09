declare interface IHelloworldWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  nameFieldLabel:string;
  stateFieldLabel:string;
}

declare module 'HelloworldWebPartStrings' {
  const strings: IHelloworldWebPartStrings;
  export = strings;
}
