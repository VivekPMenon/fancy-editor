import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { PublisherPlugin } from '../features/publisher-panel/PublisherPlugin';
import { officeAdapter } from '../adapters/officeAdapter';
import './taskpane.css';

Office.onReady(() => {
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <PublisherPlugin adapter={officeAdapter} hostLabel="Word" />
    </StrictMode>,
  );
});
