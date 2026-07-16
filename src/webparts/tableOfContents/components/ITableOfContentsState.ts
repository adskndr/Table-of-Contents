export interface ITableOfContentsState {
  historyCount: number;
  /**
   * Tracks the active tab index for every nested tab group.
   * Key = path of the tab group in the hierarchy (e.g. "root-0-1"), value = active index within that group.
   * This allows each nested level (H1 > H2 > H3 > H4) to remember its own selection independently.
   */
  activeTabPath: { [path: string]: number };
}
