import test from 'ava'
import { loadSheetProviderInDir } from "./dist/loader.js"
test('load sheet providers under a certain folder', async t => {
  const providers = await loadSheetProviderInDir("test-sheet-providers", e => {
    console.log(e)
  })
  console.log(providers)
  t.assert(providers.length === 1)
})