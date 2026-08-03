import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { PublisherPanel } from '../core/PublisherPanel';
import { officeAdapter } from '../adapters/officeAdapter';
import './taskpane.css';

Office.onReady(() => {
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <PublisherPanel adapter={officeAdapter} hostLabel="Word" />
    </StrictMode>,
  );
});
