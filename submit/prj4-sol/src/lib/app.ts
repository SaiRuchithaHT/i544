import { Result, Err, okResult } from 'cs544-js-utils';

import { Errors, makeElement } from './utils.js';

import SpreadsheetWs from './ss-ws.js';

import makeSpreadsheet from './spreadsheet.js';

export default async function makeApp(wsUrl: string) {
  makeTopLevelUI(wsUrl);
  setupLoadFormHandler();
}

function setupLoadFormHandler() {
  const errors = new Errors();
  const wsUrlInput = document.querySelector('#ws-url') as HTMLInputElement;
  const ssNameInput = document.querySelector('#ss-name') as HTMLInputElement;
  let ws: SpreadsheetWs;
  let ssName: string;
  const ssForm = document.querySelector('#ss-form')! as HTMLFormElement;
  ssForm.addEventListener('submit', async ev => {
    ev.preventDefault();
    errors.clear();
    const wsUrl = wsUrlInput.value.trim();
    const ssName = ssNameInput.value.trim();
    if (wsUrl.length === 0 || ssName.length === 0) {
      const msg =
	'both the Web Services Url and Spreadsheet Name must be specified';
      errors.display([new Err(msg, { code: 'BAD_REQ'} )]);
    }
    else {
      const ws = SpreadsheetWs.make(wsUrl);
      await makeSpreadsheet(ws, ssName);
    }
  });
}

/** Add UI corresponding to following HTML structure to #app 

  <form class="form" id="ss-form">

    <label for="ws-url">Web Services Url</label>
    <input name="ws-url" id="ws-url">

    <label for="ss-name">Spreadsheet Name</label>
    <input name="ss-name" id="ss-name">

    <label></label>
    <button type="submit">Load Spreadsheet</button>
    
  </form>

  <ul class="error" id="errors"></ul>
    
  <div id="ss">
    <!-- innerHTML of this div replaced by spreadsheet table -->
  </div>
*/
function makeTopLevelUI(wsUrl: string) {

  function makeLoadForm(wsUrl: string) {
    const form = makeElement('form', { class: 'form', id: 'ss-form' });

    form.append(makeElement('label', {for: 'ws-url'}, 'Web Services URL'));
    form.append(makeElement('input', {name: 'ws-url', id: 'ws-url',
				      value: wsUrl}));

    form.append(makeElement('label', {for: 'ss-name'}, 'Spreadsheet Name'));
    form.append(makeElement('input', {name: 'ss-name', id: 'ss-name'}));

    form.append(makeElement('label'));
    form.append(makeElement('button', {type: 'submit'}, 'Load Spreadsheet'));
    return form;
  }

  const app = document.querySelector('#app')!;
  app.append(makeLoadForm(wsUrl));

  app.append(makeElement('ul', { class: 'error', id: 'errors' }));

  //spreadsheet table should be rendered within this div
  app.append(makeElement('div', {id: 'ss'}));

}
  
