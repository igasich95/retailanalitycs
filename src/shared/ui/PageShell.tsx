import type { PropsWithChildren } from 'react'
import './page-shell.css'

export function PageShell({ children }: PropsWithChildren) {
  return (
    <div className="page-shell">
      <header className="site-header">
        <a className="brand" href="/" aria-label="Excel Table — на главную">
          <span className="brand__mark" aria-hidden="true">X</span>
          <span>Excel Table</span>
        </a>
        <span className="site-header__caption">Импорт и просмотр данных</span>
      </header>
      {children}
    </div>
  )
}

