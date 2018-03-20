import * as React from 'react';
import * as ReactDom from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import { ReactFooter } from './components/ReactFooter';
import { IReactFooterProps } from './components/IReactFooterProps';

import * as strings from 'CollabCommFooterApplicationCustomizerStrings';

const LOG_SOURCE: string = 'CollabCommFooterApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICollabCommFooterApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CollabCommFooterApplicationCustomizer
  extends BaseApplicationCustomizer<ICollabCommFooterApplicationCustomizerProperties> {

    private _footerPlaceholder: PlaceholderContent | undefined;
    
    @override
    public onInit(): Promise<void> {
      Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
  
      // Added to handle possible changes on the existence of placeholders.
      this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
      
      // Call render method for generating the HTML elements.
      this._renderPlaceHolders();
  
      return Promise.resolve();
    }
  
    @override
    private _renderPlaceHolders(): void {
  
      // Handling footer place holder
      if (!this._footerPlaceholder) {
        this._footerPlaceholder =
          this.context.placeholderProvider.tryCreateContent(
            PlaceholderName.Bottom,
            { onDispose: this._onDispose });
  
        // The extension should not assume that the expected placeholder is available.
        if (!this._footerPlaceholder) {
          console.error('The expected placeholder (Bottom) was not found.');
          return;
        }
      }
  
      const element: React.ReactElement<IReactFooterProps> = React.createElement(
        ReactFooter,
        {
          description: "The default footer"
        }
      );
  
      ReactDom.render(element, this._footerPlaceholder.domElement);
    }
  
    private _onDispose(): void {
      console.log('[CustomFooter._onDispose] Disposed custom footer.');
    }
}
