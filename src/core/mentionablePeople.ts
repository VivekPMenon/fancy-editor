export interface MentionablePerson {
  id: string;
  name: string;
  role: string;
}

// Reuses the mock post authors already established in mockPosts.ts, rather
// than inventing a separate cast of names — same demo data, one more use.
export const MENTIONABLE_PEOPLE: MentionablePerson[] = [
  { id: 'priya-nair', name: 'Priya Nair', role: 'Trade Desk' },
  { id: 'daniel-osei', name: 'Daniel Osei', role: 'Energy Desk' },
  { id: 'marta-kowalski', name: 'Marta Kowalski', role: 'Rates Desk' },
  { id: 'kenji-watanabe', name: 'Kenji Watanabe', role: 'Policy Desk' },
  { id: 'sofia-bianchi', name: 'Sofia Bianchi', role: 'Technology Desk' },
];
