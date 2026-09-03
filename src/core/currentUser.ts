// Hardcoded stand-in for the signed-in user — there's no auth in this POC.
// Drives the app header's profile chip and stamps createdBy/lastUpdatedBy on
// articles as they're created/saved (see articleStore.ts).
export interface CurrentUser {
  name: string;
  email: string;
  initials: string;
}

export const CURRENT_USER: CurrentUser = {
  name: 'Vivek Menon',
  email: 'vivek.p.menon@abc.com',
  initials: 'VM',
};
