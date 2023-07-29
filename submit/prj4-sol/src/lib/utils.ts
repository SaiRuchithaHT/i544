import { Err } from 'cs544-js-utils';

export class Errors {

  private readonly errors: HTMLElement;
  constructor() {
    this.errors = document.querySelector('#errors')!;
  }

  display(errors: Err[]) {
    errors.forEach(err => {
      this.errors.append(makeElement('li', {}, err.message));
    });
  }

  clear() { this.errors.innerHTML = ''; }

}

/** Return a new DOM element with specified tagName, attributes
 *  given by object attrs and initial contained text.
 *  Note that .append(TextOrElement) can be called on the returned
 *  element to append string text or a new DOM element to it.
 */
export function makeElement(tagName: string,
			    attrs: {[attr: string]: string} = {},
			    text='')
  : HTMLElement
{
  const element = document.createElement(tagName);
  for (const [k, v] of Object.entries(attrs)) {
    element.setAttribute(k, v);
  }
  if (text.length > 0) element.append(text);
  return element;
}
