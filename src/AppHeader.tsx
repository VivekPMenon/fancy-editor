import { useState } from 'react';
import { CURRENT_USER } from './core/currentUser';
import './AppHeader.css';

// Top-level app chrome — company brand on the left, signed-in user's quick
// profile info on the right. There's no auth in this POC, so CURRENT_USER is
// a hardcoded stand-in (see core/currentUser.ts); swapping in a real identity
// provider later only means changing that one source, not this component.
export function AppHeader() {
  const [menuOpen, setMenuOpen] = useState(false);

  return (
    <header className="app-top-header">
      <div className="app-brand">
        <img src="/favicon.svg" alt="" className="app-brand-logo" />
        <span className="app-brand-name">Publisher Plus</span>
      </div>

      <div className="app-profile">
        <button
          type="button"
          className="app-profile-trigger"
          onClick={() => setMenuOpen((open) => !open)}
          aria-haspopup="true"
          aria-expanded={menuOpen}
        >
          <span className="app-profile-avatar">{CURRENT_USER.initials}</span>
          <span className="app-profile-name">{CURRENT_USER.name}</span>
        </button>

        {menuOpen && (
          <>
            {/* Click-anywhere-outside-to-close, without wiring up a real
                outside-click listener. */}
            <button
              type="button"
              className="app-profile-backdrop"
              aria-hidden="true"
              tabIndex={-1}
              onClick={() => setMenuOpen(false)}
            />
            <div className="app-profile-menu" role="menu">
              <div className="app-profile-menu-header">
                <span className="app-profile-avatar app-profile-avatar--large">{CURRENT_USER.initials}</span>
                <div>
                  <p className="app-profile-menu-name">{CURRENT_USER.name}</p>
                  <p className="app-profile-menu-email">{CURRENT_USER.email}</p>
                </div>
              </div>
              <button type="button" className="app-profile-menu-item" disabled>
                Account settings
              </button>
              <button type="button" className="app-profile-menu-item" disabled>
                Sign out
              </button>
            </div>
          </>
        )}
      </div>
    </header>
  );
}
