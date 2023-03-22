import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import ReplyEmailComponent from './components/ReplyEmailComponent';

export default class ReplyEmailWebPart extends BaseClientSideWebPart<unknown> {

  public render(): void {
    const element: React.ReactElement<unknown> = React.createElement(ReplyEmailComponent);

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
