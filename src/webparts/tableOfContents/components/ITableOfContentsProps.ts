import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface ITableOfContentsProps {
  themeVariant: IReadonlyTheme | undefined;

  hideTitle: boolean;
  titleText: string;
  fontSize: string;

  searchText: boolean;
  searchMarkdown: boolean;
  searchCollapsible: boolean;

  showHeading2: boolean;
  showHeading3: boolean;
  showHeading4: boolean;
  showHeading5: boolean;

  showPreviousPageLinkTitle: boolean;
  showPreviousPageLinkAbove: boolean;
  showPreviousPageLinkBelow: boolean;
  previousPageText: string;

  enableStickyMode: boolean;
  webpartId: string;

  hideInMobileView: boolean;

  listStyle: string;
  isEditMode: boolean;

  layoutMode: string;
  tileBackgroundColor: string;
  tileTextColor: string;
}