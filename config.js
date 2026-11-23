// ─── M&E BIG DAY settings ────────────────────────────────────────────────
// After registering the app in Azure (see SETUP.md), paste its
// "Application (client) ID" below, commit, and you're done.
export const CONFIG = {
  clientId: 'PASTE-YOUR-CLIENT-ID-HERE',

  // 'common' lets both personal Microsoft accounts and work accounts sign in.
  authority: 'common',

  // Where the workbook lives inside the owner's OneDrive (path from the OneDrive root).
  ownerFilePath: 'Edmon Drive/TOUTA/_BIG DAY/_Wedding_Budget_Planner.xlsx',
  fileName: '_Wedding_Budget_Planner.xlsx',

  // Free, key-less USD exchange-rate feed used by "Get live rate".
  fxUrl: 'https://open.er-api.com/v6/latest/USD',
};
