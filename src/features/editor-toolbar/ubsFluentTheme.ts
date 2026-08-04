import { createLightTheme, type BrandVariants, type Theme } from '@fluentui/react-components';

// 16-step ramp Fluent's theme builder requires, anchored at shade 80 = the
// UBS red already used everywhere else in this app (buttons, chips, links).
const ubsRed: BrandVariants = {
  10: '#FDF3F3',
  20: '#FCE7E6',
  30: '#FAD0CE',
  40: '#F7B4B1',
  50: '#F39490',
  60: '#EE726C',
  70: '#E94E46',
  80: '#E60000',
  90: '#CC0000',
  100: '#B30000',
  110: '#990000',
  120: '#800000',
  130: '#660000',
  140: '#4D0000',
  150: '#330000',
  160: '#1A0000',
};

export const ubsFluentTheme: Theme = createLightTheme(ubsRed);
