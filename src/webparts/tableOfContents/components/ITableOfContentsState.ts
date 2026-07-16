export interface ITableOfContentsState {
  historyCount: number;
  /**
   * Tracks the active tab index for every nested tab group (layout = "tabs").
   * Key = path of the tab group in the hierarchy (e.g. "root-0-1"), value = active index within that group.
   */
  activeTabPath: { [path: string]: number };
  /**
   * Tracks expand/collapse state for every card, keyed by its path in the hierarchy (layout = "cards").
   * Missing entry = expanded (default).
   */
  expandedPaths: { [path: string]: boolean };
  /**
   * Custom display order for the top-level (H1) cards, as an ordered list of link keys (their text).
   * Populated from localStorage so a visitor's drag-and-drop reordering persists across visits.
   */
  cardOrderKeys: string[];
}
