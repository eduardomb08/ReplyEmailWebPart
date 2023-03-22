import * as React from 'react';

//import * as Office from 'office-js';

//// <reference types="@types/office-js" />

/* global Office */

export default class ReplyEmailComponent extends React.Component<unknown, {}> {

  public componentDidMount(): void {
    debugger;
    Office.initialize = () => {
      console.log('Office.js initialized');
      // Perform additional initialization tasks here
    };
  }

  private handleClick = (): void => {
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ['recipient@example.com'],
      subject: 'Test email',
      body: 'Hello, world!',
    });
  };

  public render(): React.ReactElement<unknown> {
    debugger;
    return (
      <section>
        <div>
          <script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js" />
          <button onClick={this.handleClick}>Send Email</button>
        </div>
      </section>
    );
  }
}
