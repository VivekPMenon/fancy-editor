export interface FlaggedTerm {
  term: string;
  comment: string;
}

/** Hardcoded compliance-style term list — POC stand-in for a real content-scanning service. */
export const FLAGGED_TERMS: FlaggedTerm[] = [
  { term: 'guarantee', comment: 'Flagged: avoid promissory language like "guarantee" in research content.' },
  { term: 'insider', comment: 'Flagged: potential compliance concern — review use of "insider".' },
  { term: 'confidential', comment: 'Flagged: check whether this content should be marked confidential.' },
  { term: 'stupid', comment: 'Flagged: unprofessional language — consider rewording.' },
];
