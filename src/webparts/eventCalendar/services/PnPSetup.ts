import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import { SPFI } from '@pnp/sp';

let _sp: SPFI | undefined;

export function initPnP(context: WebPartContext): void {
  _sp = spfi().using(SPFx(context));
}

export function getSP(): SPFI {
  if (!_sp) {
    throw new Error('PnPjs not initialized. Call initPnP(context) in onInit().');
  }
  return _sp;
}
