import { FileImport } from '../features/file-import/ui/FileImport'
import { PageShell } from '../shared/ui/PageShell'
import './app.css'

export function App() {
  return (
    <PageShell>
      <main className="workspace">
        <FileImport />
      </main>
    </PageShell>
  )
}
