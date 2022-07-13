
export interface ISourceProps {

  webUrl: string;
  listTitle: string;
  webRelativeLink: string;
  viewItemLink?: string;
  columns: string[];
  searchProps: string[];
  selectThese?: string[];
  restFilter?: string;
  searchSourceDesc: string;
  itemFetchCol?: string[]; //higher cost columns to fetch on opening panel
  orderBy?: {
      prop: string;
      asc: boolean;
  };
  defSearchButtons: string[];  //These are default buttons always on that source page.  Use case for Manual:  Policy, Instruction etc...

}


export interface IFastFiltersProps {
  description: string;
  isDarkTheme: boolean;

  sourceProps:  ISourceProps;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  
}
