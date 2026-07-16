import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneToggle,
  PropertyPaneTextField,
  PropertyPaneLabel,
  PropertyPaneDropdown
} from "@microsoft/sp-property-pane";
import { PropertyFieldColorPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';
import { PropertyFieldFilePicker, IPropertyFieldFilePickerProps } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme
} from '@microsoft/sp-component-base';

import * as strings from 'TableOfContentsWebPartStrings';
import TableOfContents from './components/TableOfContents';
import { ITableOfContentsProps, ILevelStyle } from './components/ITableOfContentsProps';

/**
 * Flat (property-pane friendly) representation of one heading level's style, stored directly
 * on the webpart properties as h1..., h2..., h3..., h4... prefixed fields.
 */
interface ILevelStyleProps {
  useCustomColors: boolean;
  backgroundColor: string;
  textColor: string;
  iconType: 'none' | 'icon' | 'image';
  iconName?: string;
  iconUrl?: string;
}

export interface ITableOfContentsWebPartProps {
  hideTitle: boolean;
  titleText: string;
  searchText: boolean;
  searchMarkdown: boolean;
  searchCollapsible: boolean;
  showHeading1: boolean;
  showHeading2: boolean;
  showHeading3: boolean;
  showHeading4: boolean;
  showPreviousPageLinkTitle: boolean;
  showPreviousPageLinkAbove: boolean;
  showPreviousPageLinkBelow: boolean;
  previousPageText: string;
  historyCount: number;
  enableStickyMode: boolean;
  hideInMobileView: boolean;
  listStyle: string;
  fontSize: string;
  layoutMode: string;
  allowCardReordering: boolean;

  h1UseCustomColors: boolean;
  h1BackgroundColor: string;
  h1TextColor: string;
  h1IconType: 'none' | 'icon' | 'image';
  h1IconName: string;
  h1IconUrl: string;

  h2UseCustomColors: boolean;
  h2BackgroundColor: string;
  h2TextColor: string;
  h2IconType: 'none' | 'icon' | 'image';
  h2IconName: string;
  h2IconUrl: string;

  h3UseCustomColors: boolean;
  h3BackgroundColor: string;
  h3TextColor: string;
  h3IconType: 'none' | 'icon' | 'image';
  h3IconName: string;
  h3IconUrl: string;

  h4UseCustomColors: boolean;
  h4BackgroundColor: string;
  h4TextColor: string;
  h4IconType: 'none' | 'icon' | 'image';
  h4IconName: string;
  h4IconUrl: string;
}

export default class TableOfContentsWebPart extends BaseClientSideWebPart<ITableOfContentsWebPartProps> {

  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  protected onInit(): Promise<void> {
    // Consume the ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();
    this.setCSSVariables(this._themeVariant.semanticColors);
    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
    // return super.onInit()
    return super.onInit().then(_ => {
      if (this.properties.searchText === undefined) {
        this.properties.searchText = true;
        this.properties.showHeading4 = true;
      }
      if (this.properties.layoutMode === undefined) {
        this.properties.layoutMode = 'list';
      }
      if (this.properties.allowCardReordering === undefined) {
        this.properties.allowCardReordering = true;
      }
      // Default styling for each level: follow the SharePoint design (no custom colors, no icon)
      // until the user explicitly opts in to custom colors/icons for that level.
      for (const level of [1, 2, 3, 4]) {
        this.setLevelStyleDefault(level);
      }
    });
  }

  /**
   * Ensures a heading level has default style property values the first time the webpart runs.
   */
  private setLevelStyleDefault(level: number): void {
    const props = this.properties as unknown as { [key: string]: unknown };
    if (props[`h${level}UseCustomColors`] === undefined) {
      props[`h${level}UseCustomColors`] = false;
    }
    if (props[`h${level}BackgroundColor`] === undefined) {
      props[`h${level}BackgroundColor`] = '#0078D4';
    }
    if (props[`h${level}TextColor`] === undefined) {
      props[`h${level}TextColor`] = '#FFFFFF';
    }
    if (props[`h${level}IconType`] === undefined) {
      props[`h${level}IconType`] = 'none';
    }
  }

  /**
   * Reads the flat, per-level property-pane fields (h1..., h2..., h3..., h4...) and assembles
   * them into the ILevelStyle[] array consumed by the React component.
   */
  private getLevelStyles(): ILevelStyle[] {
    const props = this.properties as unknown as { [key: string]: unknown };
    return [1, 2, 3, 4].map((level) => {
      const style: ILevelStyleProps = {
        useCustomColors: !!props[`h${level}UseCustomColors`],
        backgroundColor: (props[`h${level}BackgroundColor`] as string) || '#0078D4',
        textColor: (props[`h${level}TextColor`] as string) || '#FFFFFF',
        iconType: (props[`h${level}IconType`] as 'none' | 'icon' | 'image') || 'none',
        iconName: props[`h${level}IconName`] as string,
        iconUrl: props[`h${level}IconUrl`] as string
      };
      return style;
    });
  }

  private setCSSVariables(theming: any): any {
    if (!theming) { return null; }
    let themingKeys = Object.keys(theming);
    if (themingKeys !== null) {
      themingKeys.forEach(key => {
        this.domElement.style.setProperty(`--${key}`, theming[key]);
      });
    }
  }

  /**
 * Update the current theme variant reference and re-render.
 *
 * @param args The new theme
 */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.setCSSVariables(this._themeVariant.semanticColors);
    this.render();
  }

  public render(): void {
    const element: React.ReactElement<ITableOfContentsProps> = React.createElement(
      TableOfContents,
      {
        themeVariant: this._themeVariant,

        hideTitle: this.properties.hideTitle,
        titleText: this.properties.titleText,
		fontSize: this.properties.fontSize || '15px',

        searchText: this.properties.searchText,
        searchMarkdown: this.properties.searchMarkdown,
        searchCollapsible: this.properties.searchCollapsible,

        showHeading2: this.properties.showHeading1,
        showHeading3: this.properties.showHeading2,
        showHeading4: this.properties.showHeading3,
        showHeading5: this.properties.showHeading4,

        showPreviousPageLinkTitle: this.properties.showPreviousPageLinkTitle,
        showPreviousPageLinkAbove: this.properties.showPreviousPageLinkAbove,
        showPreviousPageLinkBelow: this.properties.showPreviousPageLinkBelow,
        previousPageText: this.properties.previousPageText,

        enableStickyMode: this.properties.enableStickyMode,
        webpartId: this.context.instanceId,

        hideInMobileView: this.properties.hideInMobileView,

        listStyle: this.properties.listStyle,
        isEditMode: this.displayMode == DisplayMode.Edit,

        layoutMode: this.properties.layoutMode || 'list',
        allowCardReordering: this.properties.allowCardReordering !== false,
        levelStyles: this.getLevelStyles(),
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let showHeading4: any;
    let showPreviousPageLinkTitle: any;

    if (this.properties.searchMarkdown) {
      showHeading4 = PropertyPaneCheckbox('showHeading4', {
        text: strings.showHeading4FieldLabel
      })
    }
    else {
      showHeading4 = PropertyPaneCheckbox('showHeading4', {
        text: strings.showHeading4FieldLabel,
        disabled: true
      });
    }

    if (this.properties.hideTitle) {
      showPreviousPageLinkTitle = PropertyPaneCheckbox('showPreviousPageLinkTitle', {
        text: strings.showPreviousPageTitleLabel,
        disabled: true
      })
    }
    else {
      showPreviousPageLinkTitle = PropertyPaneCheckbox('showPreviousPageLinkTitle', {
        text: strings.showPreviousPageTitleLabel
      });
    }

    return {
      pages: [
        {
          header: {
            description: strings.propertyPaneDescription

          },
          groups: [
            {
              groupFields: [
                PropertyPaneToggle('hideTitle', {
                  label: strings.hideTitleFieldLabel
                }),
                PropertyPaneTextField('titleText', {
                  description: strings.titleFieldDescription,
                  disabled: this.properties.hideTitle,
                  onGetErrorMessage: this.checkToggleField,
                  value: strings.titleDefaultValue
                }),
              ]
            },
            {
              groupFields: [
                PropertyPaneLabel('searchWebpartsLabel', {
                  text: strings.searchWebpartsLabel
                }),
                PropertyPaneCheckbox('searchText', {
                  text: strings.searchText,
                }),
                PropertyPaneCheckbox('searchMarkdown', {
                  text: strings.searchMarkdown
                }),
                PropertyPaneCheckbox('searchCollapsible', {
                  text: strings.searchCollapsible
                }),
              ]
            },
            {
              groupFields: [
                PropertyPaneLabel('showHeadingLevelsLabel', {
                  text: strings.showHeadingLevelsLabel
                }),
                PropertyPaneCheckbox('showHeading1', {
                  text: strings.showHeading1FieldLabel
                }),
                PropertyPaneCheckbox('showHeading2', {
                  text: strings.showHeading2FieldLabel
                }),
                PropertyPaneCheckbox('showHeading3', {
                  text: strings.showHeading3FieldLabel
                }),
                showHeading4,
                PropertyPaneDropdown('listStyle', {
                  label: strings.listStyle,
                  options: [
                    { key: 'default', text: 'Default' },
                    { key: 'disc', text: 'Disc' },
                    { key: 'circle', text: 'Circle' },
                    { key: 'square', text: 'Square' },
                    { key: 'none', text: 'None' }
                  ],
                  selectedKey: "default"
                }),
              ]
            },
            {
              groupFields: [
                PropertyPaneLabel('layoutModeLabel', {
                  text: 'Layout'
                }),
                PropertyPaneDropdown('layoutMode', {
                  label: 'Anzeigemodus',
                  options: [
                    { key: 'list', text: 'Liste' },
                    { key: 'tiles', text: 'Kacheln' },
                    { key: 'tabs', text: 'Tabs' },
                    { key: 'cards', text: 'Karten' }
                  ],
                  selectedKey: this.properties.layoutMode || 'list'
                }),
                PropertyPaneLabel('layoutModeDescription', {
                  text: 'Bei Kacheln, Tabs und Karten werden Ebene 1 bis 4 (H1-H4) ineinander verschachtelt dargestellt.'
                }),
                PropertyPaneToggle('allowCardReordering', {
                  label: 'Karten (Ebene 1) per Drag & Drop verschiebbar',
                  disabled: this.properties.layoutMode !== 'cards'
                })
              ]
            },
            ...([1, 2, 3, 4].map((level) => this.getLevelStyleGroup(level))),
            {
              groupFields: [
                PropertyPaneLabel('previousPageLabel', {
                  text: strings.showPreviousPageViewLabel
                }),
                showPreviousPageLinkTitle,
                PropertyPaneCheckbox('showPreviousPageLinkAbove', {
                  text: strings.showPreviousPageAboveLabel
                }),
                PropertyPaneCheckbox('showPreviousPageLinkBelow', {
                  text: strings.showPreviousPageBelowLabel
                }),
                PropertyPaneTextField('previousPageText', {
                  label: strings.previousPageFieldLabel,
                  disabled: (!this.properties.showPreviousPageLinkTitle || this.properties.hideTitle) && !this.properties.showPreviousPageLinkAbove && !this.properties.showPreviousPageLinkBelow,
                  onGetErrorMessage: this.checkToggleField,
                  value: strings.previousPageDefaultValue
                }),
              ]
            },
            {
              groupFields: [
                PropertyPaneToggle('enableStickyMode', {
                  label: strings.enableStickyModeLabel
                }),
                PropertyPaneLabel('enabldeStickyModeDescription', {
                  text: strings.enableStickyModeDescription
                }),
                PropertyPaneToggle('hideInMobileView', {
                  label: strings.hideInMobileViewLabel
                })
              ]
            },
			{
			  groupFields: [
                PropertyPaneTextField('fontSize', {
				  label: 'Schriftgröße (z. B. 16px, 14px oder 1.2rem)',
				  description: 'Gib die gewünschte Größe mit Einheit an.',
				  value: '15px' // Standardwert, falls nichts eingegeben wurde
				})
              ]
			}
          ]
        }
      ]
    };
  }

  /**
   * Builds the property pane group for one heading level (H1-H4): a toggle to choose between
   * "SharePoint design" and custom colors, the color pickers (only active if custom colors are on),
   * and an icon chooser that lets the user pick either a Fluent UI icon (icon library) or a custom image.
   */
  private getLevelStyleGroup(level: number): { groupFields: any[] } {
    const props = this.properties as unknown as { [key: string]: unknown };
    const prefix = `h${level}`;
    const useCustomColors = !!props[`${prefix}UseCustomColors`];
    const iconType = (props[`${prefix}IconType`] as string) || 'none';

    const filePickerProps: IPropertyFieldFilePickerProps & { key: string } = {
      context: this.context as any,
      filePickerResult: props[`${prefix}IconUrl`]
        ? ({ fileAbsoluteUrl: props[`${prefix}IconUrl`] } as any)
        : undefined,
      onSave: (fileResult: { fileAbsoluteUrl: string }) => {
        (this.properties as unknown as { [key: string]: unknown })[`${prefix}IconUrl`] = fileResult.fileAbsoluteUrl;
        this.render();
      },
      onChanged: (fileResult: { fileAbsoluteUrl: string }) => {
        (this.properties as unknown as { [key: string]: unknown })[`${prefix}IconUrl`] = fileResult.fileAbsoluteUrl;
        this.render();
      },
      buttonLabel: 'Bild auswählen',
      disabled: iconType !== 'image',
      key: `${prefix}IconUrlFieldId`
    } as any;

    return {
      groupFields: [
        PropertyPaneLabel(`${prefix}Label`, {
          text: `Ebene ${level} (H${level}) - Kacheln/Tabs/Karten`
        }),
        PropertyPaneToggle(`${prefix}UseCustomColors`, {
          label: 'Eigene Farben verwenden',
          onText: 'Eigene Farben',
          offText: 'SharePoint-Design'
        }),
        PropertyFieldColorPicker(`${prefix}BackgroundColor`, {
          label: 'Hintergrundfarbe',
          selectedColor: (props[`${prefix}BackgroundColor`] as string) || '#0078D4',
          onPropertyChange: this.onPropertyPaneFieldChanged,
          properties: this.properties,
          disableAlpha: true,
          disabled: !useCustomColors,
          key: `${prefix}BackgroundColorFieldId`
        } as any),
        PropertyFieldColorPicker(`${prefix}TextColor`, {
          label: 'Textfarbe',
          selectedColor: (props[`${prefix}TextColor`] as string) || '#FFFFFF',
          onPropertyChange: this.onPropertyPaneFieldChanged,
          properties: this.properties,
          disableAlpha: true,
          disabled: !useCustomColors,
          key: `${prefix}TextColorFieldId`
        } as any),
        PropertyPaneDropdown(`${prefix}IconType`, {
          label: 'Symbol',
          options: [
            { key: 'none', text: 'Kein Symbol' },
            { key: 'icon', text: 'Symbol-Bibliothek' },
            { key: 'image', text: 'Eigenes Bild' }
          ],
          selectedKey: iconType
        }),
        PropertyFieldIconPicker(`${prefix}IconName`, {
          label: 'Symbol auswählen',
          currentIcon: (props[`${prefix}IconName`] as string) || '',
          key: `${prefix}IconNameFieldId`,
          onSave: (iconName: string) => {
            (this.properties as unknown as { [key: string]: unknown })[`${prefix}IconName`] = iconName;
            this.render();
          },
          properties: this.properties,
          disabled: iconType !== 'icon',
          onPropertyChange: this.onPropertyPaneFieldChanged,
          buttonLabel: 'Symbol auswählen',
          renderOption: 'panel'
        } as any),
        PropertyFieldFilePicker(`${prefix}IconUrl`, filePickerProps as any)
      ]
    };
  }

  private checkToggleField = (value: string): string => {
    if (value === "") {
      return strings.errorToggleFieldEmpty;
    }
    else {
      return "";
    }
  }

}