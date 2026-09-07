import {
  ArrowTrendingRegular,
  BuildingBankRegular,
  EarthRegular,
  DocumentTextRegular,
  WeatherSunnyRegular,
  ChartMultipleRegular,
  LightbulbFilamentRegular,
  WalletRegular,
  NewsRegular,
  GridRegular,
  RocketRegular,
  HandshakeRegular,
  CompassNorthwestRegular,
  CalendarClockRegular,
  CalculatorRegular,
  type FluentIcon,
} from '@fluentui/react-icons';
import { TICKER_DATABASE } from '../tickerDatabase';

// The 15 article-section templates (matches the reference "create-article"
// tile picker: Stock, Sector, Geography, Rich Text, Premarket, Index,
// Thematic, Financing, Flow-Commentary, Beta Solutions, Equity Capital
// Markets, Corporate Access, Tactical Market View, Earnings Preview, Quant).
// One config drives everything downstream — the tile grid, the slash-menu
// entries, and each inserted section's own header control — so adding a
// 16th template is a one-entry change here, not a new node type.
export type TemplateSectionType =
  | 'stock'
  | 'sector'
  | 'geography'
  | 'richText'
  | 'premarket'
  | 'index'
  | 'thematic'
  | 'financing'
  | 'flowCommentary'
  | 'betaSolutions'
  | 'equityCapitalMarkets'
  | 'corporateAccess'
  | 'tacticalMarketView'
  | 'earningsPreview'
  | 'quant';

// 'lookup' = free-typing typeahead over a large/structured list (tickers);
// 'select' = typeahead over a short fixed list of options; 'none' = no
// selector at all (Rich Text is just a labeled freeform section).
export type SectionControlKind = 'lookup' | 'select' | 'none';

export interface TemplateSectionConfig {
  type: TemplateSectionType;
  label: string;
  icon: FluentIcon;
  control: SectionControlKind;
  /** Options to search/select from — only for control: 'select'. */
  options?: string[];
  controlPlaceholder?: string;
  bodyPlaceholder: string;
}

// Curated, not exhaustive — same spirit as tickerDatabase.ts/flaggedTerms.ts
// elsewhere in this app: enough real-looking values to demo convincingly,
// trivially extendable later.
const SECTORS = [
  'Technology',
  'Financials',
  'Energy',
  'Industrials',
  'Healthcare',
  'Consumer Discretionary',
  'Consumer Staples',
  'Materials',
  'Utilities',
  'Real Estate',
  'Communication Services',
];

const COUNTRIES = [
  'United States',
  'United Kingdom',
  'Germany',
  'France',
  'Italy',
  'Spain',
  'Switzerland',
  'Netherlands',
  'Sweden',
  'Norway',
  'Denmark',
  'Ireland',
  'Belgium',
  'Austria',
  'Finland',
  'Portugal',
  'Greece',
  'Poland',
  'Czech Republic',
  'Luxembourg',
  'Japan',
  'China',
  'India',
  'South Korea',
  'Taiwan',
  'Hong Kong',
  'Singapore',
  'Indonesia',
  'Thailand',
  'Vietnam',
  'Malaysia',
  'Philippines',
  'Australia',
  'New Zealand',
  'Canada',
  'Brazil',
  'Mexico',
  'Chile',
  'Argentina',
  'Colombia',
  'South Africa',
  'Nigeria',
  'Egypt',
  'UAE',
  'Saudi Arabia',
  'Qatar',
  'Kuwait',
  'Israel',
  'Turkey',
  'Russia',
];

const MARKET_SESSIONS = ['US Premarket', 'European Open', 'APAC Close', 'US After-Hours'];

const INDICES = [
  'S&P 500',
  'Nasdaq 100',
  'Dow Jones Industrial Average',
  'Russell 2000',
  'FTSE 100',
  'DAX',
  'CAC 40',
  'Euro Stoxx 50',
  'Nikkei 225',
  'Hang Seng',
  'Shanghai Composite',
  'MSCI Emerging Markets',
];

const THEMES = [
  'AI Infrastructure',
  'Energy Transition',
  'Reshoring & Supply Chains',
  'Aging Demographics',
  'Digital Payments',
  'Cybersecurity',
  'Electrification',
  'Biotech Innovation',
];

const FINANCING_TYPES = [
  'IPO',
  'Follow-on Offering',
  'Convertible Bond',
  'High Yield Bond',
  'Investment Grade Bond',
  'Leveraged Loan',
  'Private Placement',
];

const ASSET_CLASSES = ['Equities', 'Rates', 'FX', 'Credit', 'Commodities'];

const BETA_SOLUTIONS = ['Index Replication', 'Smart Beta', 'Factor Investing', 'Structured Products'];

const ECM_DEAL_TYPES = ['IPO', 'Follow-on', 'Block Trade', 'Convertible', 'Accelerated Bookbuild'];

const CORPORATE_ACCESS_EVENTS = ['Roadshow', 'Investor Conference', 'Analyst Day', 'Earnings Call', 'Site Visit'];

const TACTICAL_HORIZONS = ['Short-term (0-3m)', 'Medium-term (3-12m)', 'Long-term (12m+)'];

const QUANT_FACTORS = ['Momentum', 'Value', 'Quality', 'Low Volatility', 'Growth', 'Size'];

const TICKER_OPTIONS = TICKER_DATABASE.map((t) => `${t.symbol} — ${t.name}`);

export const TEMPLATE_SECTION_TYPES: TemplateSectionConfig[] = [
  {
    type: 'stock',
    label: 'Stock',
    icon: ArrowTrendingRegular,
    control: 'lookup',
    options: TICKER_OPTIONS,
    controlPlaceholder: 'Search ticker or company…',
    bodyPlaceholder: 'Write your stock commentary…',
  },
  {
    type: 'sector',
    label: 'Sector',
    icon: BuildingBankRegular,
    control: 'select',
    options: SECTORS,
    controlPlaceholder: 'Select a sector…',
    bodyPlaceholder: 'Write your sector commentary…',
  },
  {
    type: 'geography',
    label: 'Geography',
    icon: EarthRegular,
    control: 'select',
    options: COUNTRIES,
    controlPlaceholder: 'Select a country…',
    bodyPlaceholder: 'Write your geography commentary…',
  },
  {
    type: 'richText',
    label: 'Rich Text',
    icon: DocumentTextRegular,
    control: 'none',
    bodyPlaceholder: 'Write freeform content…',
  },
  {
    type: 'premarket',
    label: 'Premarket',
    icon: WeatherSunnyRegular,
    control: 'select',
    options: MARKET_SESSIONS,
    controlPlaceholder: 'Select a session…',
    bodyPlaceholder: 'Write your premarket commentary…',
  },
  {
    type: 'index',
    label: 'Index',
    icon: ChartMultipleRegular,
    control: 'select',
    options: INDICES,
    controlPlaceholder: 'Select an index…',
    bodyPlaceholder: 'Write your index commentary…',
  },
  {
    type: 'thematic',
    label: 'Thematic',
    icon: LightbulbFilamentRegular,
    control: 'select',
    options: THEMES,
    controlPlaceholder: 'Select a theme…',
    bodyPlaceholder: 'Write your thematic commentary…',
  },
  {
    type: 'financing',
    label: 'Financing',
    icon: WalletRegular,
    control: 'select',
    options: FINANCING_TYPES,
    controlPlaceholder: 'Select a financing type…',
    bodyPlaceholder: 'Write your financing commentary…',
  },
  {
    type: 'flowCommentary',
    label: 'Flow-Commentary',
    icon: NewsRegular,
    control: 'select',
    options: ASSET_CLASSES,
    controlPlaceholder: 'Select an asset class…',
    bodyPlaceholder: 'Write your flow commentary…',
  },
  {
    type: 'betaSolutions',
    label: 'Beta Solutions',
    icon: GridRegular,
    control: 'select',
    options: BETA_SOLUTIONS,
    controlPlaceholder: 'Select a solution…',
    bodyPlaceholder: 'Write your beta solutions commentary…',
  },
  {
    type: 'equityCapitalMarkets',
    label: 'Equity Capital Markets',
    icon: RocketRegular,
    control: 'select',
    options: ECM_DEAL_TYPES,
    controlPlaceholder: 'Select a deal type…',
    bodyPlaceholder: 'Write your ECM commentary…',
  },
  {
    type: 'corporateAccess',
    label: 'Corporate Access',
    icon: HandshakeRegular,
    control: 'select',
    options: CORPORATE_ACCESS_EVENTS,
    controlPlaceholder: 'Select an event type…',
    bodyPlaceholder: 'Write your corporate access commentary…',
  },
  {
    type: 'tacticalMarketView',
    label: 'Tactical Market View',
    icon: CompassNorthwestRegular,
    control: 'select',
    options: TACTICAL_HORIZONS,
    controlPlaceholder: 'Select a horizon…',
    bodyPlaceholder: 'Write your tactical view…',
  },
  {
    type: 'earningsPreview',
    label: 'Earnings Preview',
    icon: CalendarClockRegular,
    control: 'lookup',
    options: TICKER_OPTIONS,
    controlPlaceholder: 'Search ticker or company…',
    bodyPlaceholder: 'Write your earnings preview…',
  },
  {
    type: 'quant',
    label: 'Quant',
    icon: CalculatorRegular,
    control: 'select',
    options: QUANT_FACTORS,
    controlPlaceholder: 'Select a factor…',
    bodyPlaceholder: 'Write your quant commentary…',
  },
];

export function getTemplateSectionConfig(type: string): TemplateSectionConfig {
  return TEMPLATE_SECTION_TYPES.find((t) => t.type === type) ?? TEMPLATE_SECTION_TYPES[3]; // Rich Text fallback
}
