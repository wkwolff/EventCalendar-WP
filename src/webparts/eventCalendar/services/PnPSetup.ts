/**
 * @file PnPSetup.ts
 * @description Singleton initializer for PnPjs (SharePoint Framework helper library).
 *              Must be called once during `onInit()` of the web part so that all
 *              service modules can retrieve a pre-configured SPFI instance.
 * @author W. Kevin Wolff
 * @copyright TidalHealth
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import { SPFI } from '@pnp/sp';

/**
 * Module-level singleton — holds the configured PnPjs SPFI instance.
 * Remains `undefined` until {@link initPnP} is called.
 */
let _sp: SPFI | undefined;

/**
 * Initializes the PnPjs SPFI factory with the current SPFx web part context.
 *
 * Call this **exactly once** inside the web part's `onInit()` method, before
 * any React rendering occurs. Subsequent calls will overwrite the instance.
 *
 * @param context - The SPFx {@link WebPartContext} provided by the framework.
 *
 * @example
 * ```ts
 * protected async onInit(): Promise<void> {
 *   await super.onInit();
 *   initPnP(this.context);
 * }
 * ```
 */
export function initPnP(context: WebPartContext): void {
  _sp = spfi().using(SPFx(context));
}

/**
 * Returns the initialized PnPjs SPFI instance.
 *
 * @returns The singleton SPFI instance configured with the SPFx context.
 * @throws {Error} If called before {@link initPnP} has been invoked.
 */
export function getSP(): SPFI {
  if (!_sp) {
    throw new Error('PnPjs not initialized. Call initPnP(context) in onInit().');
  }
  return _sp;
}
