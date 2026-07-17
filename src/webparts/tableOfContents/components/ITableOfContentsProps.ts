/**
 * Styling configuration for one heading level (H1-H4) used by the "tiles" and "cards" layouts.
 */
export interface ILevelStyle {
  /**
   * If true, backgroundColor/textColor below are used.
   * If false, the level follows the SharePoint page theme (SharePoint design).
   */
  useCustomColors: boolean;
  backgroundColor: string;
  textColor: string;
  /**
   * 'none'   -> no icon
   * 'icon'   -> Fluent UI icon library (icon library, chosen via iconName)
   * 'image'  -> custom image (chosen via iconUrl, e.g. from a SharePoint library)
   */
  iconType: 'none' | 'icon' | 'image';
  iconName?: string;
  iconUrl?: string;
}

export interface ITableOfContentsProps {
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
  /**
   * Styling (colors + icon) for each heading level, index 0 = Level 1 (H1) ... index 3 = Level 4 (H4).
   */
  levelStyles: ILevelStyle[];
}