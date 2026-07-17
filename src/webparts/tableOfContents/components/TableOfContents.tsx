import * as React from 'react';
import styles from './TableOfContents.module.scss';
import { ITableOfContentsProps, ILevelStyle } from './ITableOfContentsProps';
import { ITableOfContentsState } from './ITableOfContentsState';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from "TableOfContentsWebPartStrings";
import { Icon } from '@fluentui/react/lib/Icon';

/**
 * Describes a link for a header
 */
interface Link {
  /**
   * The Source html element.
   */
  element: HTMLElement | undefined;
  /**
   * Child nodes for the link.
   */
  childNodes: Link[];
  /**
   * Parent link. Undefined for the root link.
   */
  parent: Link | undefined;
}

export default class TableOfContents extends React.Component<ITableOfContentsProps, ITableOfContentsState> {
  private static timeout = 500;

  private static h2Tag = "h2";
  private static h3Tag = "h3";
  private static h4Tag = "h4";
  private static h5Tag = "h5";

  /**
   * Fallback style used for a level if the property pane hasn't provided one yet (e.g. older saved webpart instances).
   */
  private static defaultLevelStyle: ILevelStyle = {
    useCustomColors: false,
    backgroundColor: '#0078D4',
    textColor: '#FFFFFF',
    iconType: 'none'
  };

  constructor(props: ITableOfContentsProps) {
    super(props);
    this.state = {
      historyCount: -1,
      expandedPaths: {}
    };
  }

  /**
   * Returns the style configuration (colors + icon) for a given nesting depth (0 = H1/Level 1, 1 = H2/Level 2, ...).
   */
  private getLevelStyle(depth: number): ILevelStyle {
    const levelStyles = this.props.levelStyles;
    if (levelStyles && levelStyles[depth]) {
      return levelStyles[depth];
    }
    return TableOfContents.defaultLevelStyle;
  }

  /**
   * Computes the inline style for a tile chip at a given level.
   * If the level is configured to use custom colors, those are applied directly.
   * Otherwise the chip follows the SharePoint page theme (via the CSS custom properties set from the site theme).
   */
  private getChipStyle(levelStyle: ILevelStyle, fontSize: string): React.CSSProperties {
    if (levelStyle.useCustomColors) {
      const bgColor = levelStyle.backgroundColor || TableOfContents.defaultLevelStyle.backgroundColor;
      const textColor = levelStyle.textColor || TableOfContents.defaultLevelStyle.textColor;
      return { fontSize, backgroundColor: bgColor, color: textColor };
    }

    // "SharePoint design": rely on the theme CSS variables set by the webpart from the current site theme.
    return { fontSize, backgroundColor: 'var(--primaryButtonBackground)', color: 'var(--primaryButtonText)' };
  }

  /**
   * Renders the icon for a chip at a given level: none, a Fluent UI icon (icon library), or a custom image.
   * Falls back to the default decorative icon if no level style is configured at all.
   */
  private renderLevelIcon(levelStyle: ILevelStyle): JSX.Element | null {
    if (!levelStyle || levelStyle.iconType === 'none') {
      return null;
    }
    if (levelStyle.iconType === 'image' && levelStyle.iconUrl) {
      return <img src={levelStyle.iconUrl} className={styles.chipIconImage} alt="" aria-hidden="true" />;
    }
    if (levelStyle.iconType === 'icon' && levelStyle.iconName) {
      return <Icon iconName={levelStyle.iconName} className={styles.chipIcon} aria-hidden="true" />;
    }
    return null;
  }

  /**
   * Gets a nested list of links based on the list of headers specified.
   * @param headers List of HtmlElements for H2, H3, and H4 headers.
   */
  private getLinks(headers: HTMLElement[]): Link[] {
    // create a root link that will be a root for links' tree
    const root: Link = { childNodes: [], parent: undefined, element: undefined };

    let prevLink: Link | null = null;

    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      const link: Link = { childNodes: [], parent: undefined, element: header };

      if (i === 0) {
        // the first header is always added as a child of the root
        link.parent = root;
        root.childNodes.push(link);
      } else {
        const prevHeader = headers[i - 1];

        // compare the current header and the previous one to define where to add new link
        const compare = this.compareHeaders(header.tagName, prevHeader.tagName);

        if (compare === 0) {
          // if headers are on the same level, add header to the same parent
          link.parent = prevLink.parent;
          prevLink.parent.childNodes.push(link);
        } else if (compare < 0) {

          let targetParent = prevLink.parent;
          // if current header is bigger than the previous one, go up in the hierarchy to find a place to add link
          // go up in the hierarchy of links until a link with bigger tag is found or until the root link found
          // i.e. for H4 look for H3 or H2, for H3 look for H2, for H2 look for the root.
          while ((targetParent != root) && (this.compareHeaders(header.tagName, targetParent.element.tagName) <= 0)) {
            targetParent = targetParent.parent;
          }

          link.parent = targetParent;
          targetParent.childNodes.push(link);
        } else {
          // if current header is smaller than the previous one, add link for it as a child of the previous link
          link.parent = prevLink;
          prevLink.childNodes.push(link);
        }
      }

      prevLink = link;
    }

    // return list of links for top-level headers
    return root.childNodes;
  }

  /**
   * Compares two header tags by their weights.
   * The function is used to compare the size of headers (e.g. should H3 go under H2?)
   * @param header1
   * @param header2
   */
  private compareHeaders(header1: string, header2: string): number {
    return this.getHeaderWeight(header1) - this.getHeaderWeight(header2);
  }

  /**
   * Returns a digital weight of a tag. Used for comparing header tags.
   * @param header
   */
  private getHeaderWeight(header: string): number {
    switch (header.toLowerCase()) {
      case (TableOfContents.h2Tag):
        return 2;
      case (TableOfContents.h3Tag):
        return 3;
      case (TableOfContents.h4Tag):
        return 4;
      case (TableOfContents.h5Tag):
        return 5;
      default:
        throw new Error('Unknown header: ' + header);
    }
  }

  /**
   * Returns html elements in the current page specified by the query selector.
   */
  private getHtmlElements(querySelector: string): HTMLElement[] {
    if (querySelector.length === 0) {
      return [];
    } else {
      const div = document.getElementById('spPageCanvasContent'); //this is the main content area of the page
      const elements = div.querySelectorAll(querySelector);
      const htmlElements: HTMLElement[] = [];

      for (let i = 0; i < elements.length; i++) {

        // While in edit mode Section headers are not headers, but text areas. This converts them to H2 tags
        if (elements[i].tagName === "TEXTAREA") {
          let temp = document.createElement('h2')
          temp.innerHTML = elements[i].innerHTML
          htmlElements.push(temp)
        }
        else {
          htmlElements.push(elements[i] as HTMLElement);
        }

      }
      return htmlElements;
    }
  }

  /**
   * Returns a query selector based on the specified props
   * @param props
   */
  private getQuerySelector(props: ITableOfContentsProps) {
    let queryParts = [];
    let queryItems = [];

    if (this.props.searchText) {
      queryItems.push('.cke_editable', '.ck-content');
    }

    if (this.props.searchCollapsible) {
      queryItems.push('[data-automation-id*="CanvasZone-SectionContainer"]');
    }

    if (this.props.searchMarkdown) {
      queryItems.push('[data-sp-feature-tag*="Markdown"]');
    }

    if (props.showHeading2) {
      for (let i = 0; i < queryItems.length; i++) {

        if (queryItems[i] === '[data-automation-id*="CollapsibleLayer-TitleInput"]') {
          queryParts.push(queryItems[i]);
        }
        else {
          queryParts.push(queryItems[i] + " " + TableOfContents.h2Tag);
        }
      }
    }

    if (props.showHeading3) {
      for (let i = 0; i < queryItems.length; i++) {
        queryParts.push(queryItems[i] + " " + TableOfContents.h3Tag);
      }
    }

    if (props.showHeading4) {
      for (let i = 0; i < queryItems.length; i++) {
        queryParts.push(queryItems[i] + " " + TableOfContents.h4Tag);
      }
    }

    if (props.showHeading5) {
      for (let i = 0; i < queryItems.length; i++) {
        queryParts.push(queryItems[i] + " " + TableOfContents.h5Tag);
      }
    }

    return queryParts.join(',');
  }

  /**
   * Filters elements with empty text.
   * @param element
   */
  private filterEmpty(element: HTMLElement): boolean {
    // Check if element is empty. If it is in a collapsible section with a 'Premalink' then return true as we can fix that later.
    if (element.innerText.trim() !== '') {
      return true;
    }
    else if (element.firstElementChild !== null) {
      if (element.firstElementChild.getAttribute('role') === 'link') {
        return true;
      }
      else {
        return false;
      }
    }
  }

  /**
   * Filters elements that are inside <aside> tag and thus not related to a page.
   * @param element
   */
  private filterAside(element: HTMLElement): boolean {
    let inAsideTag = false;

    let parentElement = element.parentElement;

    while (parentElement) {
      if (parentElement.tagName.toLocaleLowerCase() === 'aside') {
        inAsideTag = true;
        break;
      }

      parentElement = parentElement.parentElement;
    }

    return !inAsideTag;
  }

  /**
   * Filters elements that have the data attrribute of 'data-toc-ignore' and thus should be ignored.
   * @param element
   */
  private filterTocIgnore(element: HTMLElement): boolean {
    return !(element.getAttribute("data-toc-ignore"));
  }

  /**
   * Filters elements that have been set with a sytle of 'display: none'
   * @param element
   */
  private filterStyleDisplayNone(element: HTMLElement): boolean {
    let styleDisplayNone = false;

    let parentElement = element.parentElement;

    while (parentElement) {
      if (parentElement.style.display.toLocaleLowerCase() === 'none') {
        styleDisplayNone = true;
        break;
      }

      parentElement = parentElement.parentElement;
    }

    return !styleDisplayNone;
  }

  /**
   * Returns a click handler that scrolls a page to the specified element.
   */
  private scrollToHeader = (target: HTMLElement) => {
    return (event: React.SyntheticEvent) => {
      //decrement the history count to allow the return to previous page to work correctly
      const temp = this.state.historyCount - 1;
      this.setState({ historyCount: temp });
      event.preventDefault();
      document.location.hash = target.id;
      target.scrollIntoView({ behavior: 'smooth', block: 'start', inline: 'nearest' });
    };
  }

  /**
   * Extracts the display text for a link, falling back to the Permalink title if empty.
   * @param link
   */
  private getLinkText(link: Link): string {
    let linkText = link.element.innerText;
    const regex = /title="Permalink for ([^"]+)"/;

    if (linkText === "") {
      if (link.element.firstElementChild.getAttribute('role') === 'link') {
        const match = link.element.innerHTML.match(regex);
        if (match && match.length >= 2) {
          linkText = match[1];
        }
        else {
          linkText = 'Error!';
        }
      }
      else {
        linkText = 'Error!';
      }
    }

    return linkText;
  }


  /**
   * Creates a list of components to display from a list of links.
   * @param links
   */
  private renderLinks(links: Link[], listStyle: string): JSX.Element[] {
    // For each link render a <li> element with a link. If the link has got childNodes, additionaly render <ul> with child links.
    const elements = links.map((link, index) => {
      const linkText = this.getLinkText(link);

      // Hier wird die Schriftgröße aus den Props ausgelesen
      const customFontSize = this.props.fontSize || '18px';

      return (
        <li key={index} style={{ fontSize: customFontSize }}>
          <a
            onClick={this.scrollToHeader(link.element)}
            href={'#' + link.element.id}
            style={{ fontSize: customFontSize }}
          >
            {linkText}
          </a>
          {link.childNodes.length > 0 ? (<ul style={{ listStyleType: listStyle }}>{this.renderLinks(link.childNodes, listStyle)}</ul>) : ''}
        </li>
      );
    });

    return elements;
  }

  /**
   * Renders headers as a grid of rounded, coloured chip tiles, nested recursively so that
   * H1 > H2 > H3 > H4 each appear as their own tile group inside their parent tile.
   * @param links
   * @param listStyle
   * @param depth nesting depth: 0 = Level 1 (H1), 1 = Level 2 (H2), 2 = Level 3 (H3), 3 = Level 4 (H4)
   */
  private renderTiles(links: Link[], listStyle: string, depth: number = 0): JSX.Element {
    if (!links || links.length === 0) {
      return depth === 0 ? <div className={styles.tilesContainer} /> : null;
    }

    const levelStyle = this.getLevelStyle(depth);
    const customFontSize = this.props.fontSize || '18px';
    const chipStyle = this.getChipStyle(levelStyle, customFontSize);
    const containerClass = depth === 0 ? styles.tilesContainer : styles.tilesContainerNested;

    return (
      <div className={containerClass}>
        {links.map((link, index) => {
          const linkText = this.getLinkText(link);

          return (
            <div className={styles.tile} key={index}>
              <a
                className={styles.tileChip}
                onClick={this.scrollToHeader(link.element)}
                href={'#' + link.element.id}
                style={chipStyle}
              >
                {this.renderLevelIcon(levelStyle)}
                <span>{linkText}</span>
              </a>
              {link.childNodes.length > 0 ? this.renderTiles(link.childNodes, listStyle, depth + 1) : null}
            </div>
          );
        })}
      </div>
    );
  }

  /**
   * Force the component to re-render with a specified interval.
   * This is needed to get valid id values for headers to use in links. Right after the rendering headers won't have valid ids, they are assigned later once the whole page got rendered.
   * The component will display the correct list of headers on the first render and will be able to process clicks (as a link to an HTMLElement is stored by the component).
   * Once valid ids got assigned to headers by SharePoint code, the component will get valid ids for headers. This way a link from ToC can be copied by a user and it will be a valid link to a header.
   */
  public componentDidMount() {
    setInterval(() => {
      this.setState({});
    }, TableOfContents.timeout);
  }

  /**
   * Counts every descendant link nested under the given link (children, grandchildren, ...).
   * Used for the small count badge shown on a card (layout = "cards").
   */
  private countDescendants(link: Link): number {
    let count = 0;
    for (const child of link.childNodes) {
      count += 1 + this.countDescendants(child);
    }
    return count;
  }

  /**
   * Whether the card at the given path is currently expanded. Defaults to expanded.
   */
  private isCardExpanded(path: string): boolean {
    return this.state.expandedPaths[path] !== false;
  }

  /**
   * Toggles expand/collapse for the card at the given path.
   */
  private toggleCardExpanded = (path: string) => {
    return (event: React.SyntheticEvent) => {
      event.preventDefault();
      event.stopPropagation();
      const isExpanded = this.isCardExpanded(path);
      this.setState((prevState) => ({
        expandedPaths: { ...prevState.expandedPaths, [path]: !isExpanded }
      }));
    };
  }

  /**
   * Small chevron icon indicating expand/collapse state of a card.
   */
  private renderChevron(isExpanded: boolean): JSX.Element {
    return (
      <svg
        className={isExpanded ? styles.cardChevronExpanded : styles.cardChevron}
        viewBox="0 0 12 12"
        width="12"
        height="12"
        aria-hidden="true"
        focusable="false"
      >
        <path d="M2 4l4 4 4-4" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    );
  }

  /**
   * Computes the full colour set for a card at a given level: background, text, badge and
   * divider colours - all as one consistent, always-legible pair.
   *
   * Uses the exact same background/text colors as the Kacheln chips - either the level's
   * custom colors, or (if not customized) the same theme token pair the chips already use.
   */
  private getCardChrome(levelStyle: ILevelStyle): {
    cardStyle: React.CSSProperties;
    badgeStyle: React.CSSProperties;
    dividerStyle: React.CSSProperties;
  } {
    const bg = levelStyle.useCustomColors
      ? (levelStyle.backgroundColor || TableOfContents.defaultLevelStyle.backgroundColor)
      : 'var(--primaryButtonBackground)';
    const text = levelStyle.useCustomColors
      ? (levelStyle.textColor || TableOfContents.defaultLevelStyle.textColor)
      : 'var(--primaryButtonText)';

    return {
      cardStyle: { backgroundColor: bg, color: text, borderLeftColor: bg },
      badgeStyle: { backgroundColor: 'rgba(255, 255, 255, 0.25)', color: text },
      dividerStyle: { borderTopColor: 'rgba(255, 255, 255, 0.3)' }
    };
  }

  /**
   * Renders headers as collapsible "cards" (icon, title, count badge, coloured accent), nested
   * recursively so that H1 > H2 > H3 > H4 each appear as their own card indented inside their
   * parent's card - i.e. every level is its own card, but stays visually subordinate to its parent.
   * @param links
   * @param listStyle
   * @param depth nesting depth: 0 = Level 1 (H1), 1 = Level 2 (H2), 2 = Level 3 (H3), 3 = Level 4 (H4)
   * @param path identifies this card's position in the hierarchy, used as a key into expandedPaths
   */
  private renderCards(links: Link[], listStyle: string, depth: number = 0, path: string = 'root'): JSX.Element {
    if (!links || links.length === 0) {
      return depth === 0 ? <div className={styles.cardsGrid} /> : null;
    }

    const levelStyle = this.getLevelStyle(depth);
    const isTopLevel = depth === 0;
    const containerClass = isTopLevel ? styles.cardsGrid : styles.cardsNestedGroup;
    const chrome = this.getCardChrome(levelStyle);
    const customFontSize = this.props.fontSize || '18px';

    return (
      <div className={containerClass}>
        {links.map((link, index) => {
          const linkText = this.getLinkText(link);
          const cardPath = `${path}-${index}`;
          const hasChildren = link.childNodes.length > 0;
          const expanded = this.isCardExpanded(cardPath);
          const descendantCount = this.countDescendants(link);
          const icon = this.renderLevelIcon(levelStyle);

          return (
            <div
              className={styles.card}
              key={linkText + index}
              style={chrome.cardStyle}
            >
              <a
                className={styles.cardHeader}
                href={'#' + link.element.id}
                onClick={this.scrollToHeader(link.element)}
              >
                {icon ? <span className={styles.cardIcon}>{icon}</span> : null}
                <span className={styles.cardTitle} style={{ fontSize: customFontSize }}>{linkText}</span>
                {descendantCount > 0 ? <span className={styles.cardBadge} style={chrome.badgeStyle}>{descendantCount}</span> : null}
                {hasChildren ? (
                  <button
                    type="button"
                    className={styles.cardToggle}
                    aria-expanded={expanded}
                    aria-label={expanded ? 'Zusammenklappen' : 'Ausklappen'}
                    onClick={this.toggleCardExpanded(cardPath)}
                  >
                    {this.renderChevron(expanded)}
                  </button>
                ) : null}
              </a>
              {hasChildren && expanded ? (
                <React.Fragment>
                  <div className={styles.cardDivider} style={chrome.dividerStyle} />
                  {this.renderCards(link.childNodes, listStyle, depth + 1, cardPath)}
                </React.Fragment>
              ) : null}
            </div>
          );
        })}
      </div>
    );
  }

  /**
   * Event for the back to previous page link.
   * It uses the history count to work out how many pages to go back, as each click to a header results in history
   */
  public backToPreviousPage() {
    window.history.go(this.state.historyCount);
  }

  /**
   * Render the back to previous link
   */
  private renderBackToPreviousLink = (listStyle: string): JSX.Element => {
    return (
      <div className={styles.backItem} ><ul style={{ listStyleType: listStyle }}><li><a href="#" onClick={() => this.backToPreviousPage()}>{this.props.previousPageText ? this.props.previousPageText : strings.previousPageDefaultValue}</a></li></ul></div>
    );
  }

  /**
   * Modify the CSS of the appropriate HTML elements based on the wepart ID to enable sticky mode.
   * This does involve modifying HTML elements outside of the webpart, so may well break in the furture if Microsoft change their HTML\CSS etc.
   */
  private configureSticky() {

    const HTMLElementSticky: HTMLElement = document.querySelector("[id='" + this.props.webpartId + "']");
    if (HTMLElementSticky != null) {
      if (this.props.enableStickyMode && window.innerWidth > 1024) {

        if (this.props.isEditMode){
          HTMLElementSticky.parentElement.parentElement.style.position = "Sticky";
          HTMLElementSticky.parentElement.parentElement.style.top = "0px";
          HTMLElementSticky.parentElement.parentElement.parentElement.style.height = "100%";
          console.log("Edit Mode");
        }
        else {
          HTMLElementSticky.style.position = "Sticky";
          HTMLElementSticky.style.top = "0px";
          HTMLElementSticky.parentElement.style.height = "100%";
          console.log("Normal Mode");
        }
      }
      else {
        HTMLElementSticky.style.position = "";
        HTMLElementSticky.style.top = "";
        HTMLElementSticky.parentElement.style.height = "";
        HTMLElementSticky.parentElement.parentElement.style.position = "";
        HTMLElementSticky.parentElement.parentElement.style.top = "";
        HTMLElementSticky.parentElement.parentElement.parentElement.style.height = "";
      }
    }
  }

  public render(): JSX.Element {
    // get headers, then filter out empty and headers from <aside> tags
    const listStyle = escape(this.props.listStyle) === "default" ? "" : this.props.listStyle;
    const querySelector = this.getQuerySelector(this.props);
    const headers = this.getHtmlElements(querySelector).filter(this.filterEmpty).filter(this.filterAside).filter(this.filterTocIgnore).filter(this.filterStyleDisplayNone);
    // create a list of links from headers
    const links = this.getLinks(headers);

    // create components from a list of links, depending on the selected layout mode
    let toc: JSX.Element;
    switch (this.props.layoutMode) {
      case 'tiles':
        toc = this.renderTiles(links, listStyle);
        break;
      case 'cards':
        toc = this.renderCards(links, listStyle);
        break;
      case 'list':
      default:
        toc = (<ul style={{ listStyleType: listStyle }}>{this.renderLinks(links, listStyle)}</ul>);
        break;
    }

    // create previous page link
    const previousPageTitle = this.props.showPreviousPageLinkTitle && !this.props.hideTitle ? (this.renderBackToPreviousLink(listStyle)) : null;
    const previousPageAbove = this.props.showPreviousPageLinkAbove ? (this.renderBackToPreviousLink(listStyle)) : null;
    const previousPageBelow = this.props.showPreviousPageLinkBelow ? (this.renderBackToPreviousLink(listStyle)) : null;
    // add CSS class to hide in mobile view if needed
    const hideInMobileViewClass = this.props.hideInMobileView ? (styles.hideInMobileView) : '';
    // add CSS class to hide title if requested
    const titleClass = this.props.hideTitle ? (styles.hideTitle) : "cke_editable h2 " + styles.title;
    // set title text
    const titleText = this.props.titleText ? this.props.titleText : strings.titleDefaultValue;
    // set Sticky parameters
    this.configureSticky();

    return (
      <section className={styles.tableOfContents}>
        <div className={hideInMobileViewClass}>
          <nav>
            {previousPageTitle}
            <div className={titleClass}>
              <h2 data-toc-ignore="true">{titleText}</h2>
            </div>
            {previousPageAbove}
            {toc}
            {previousPageBelow}
          </nav>
        </div>
      </section>
    );
  }
}
