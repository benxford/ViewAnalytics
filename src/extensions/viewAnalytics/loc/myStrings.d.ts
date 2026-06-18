declare interface IViewAnalyticsCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ViewAnalyticsCommandSetStrings' {
  const strings: IViewAnalyticsCommandSetStrings;
  export = strings;
}
