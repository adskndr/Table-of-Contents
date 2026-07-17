export interface ITableOfContentsState {
  historyCount: number;
  /**
   * Tracks expand/collapse state for every card, keyed by its path in the hierarchy (layout = "cards").
   * Missing entry = expanded (default).
   */
  expandedPaths: { [path: string]: boolean };
}
