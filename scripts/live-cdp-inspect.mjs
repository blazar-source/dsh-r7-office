/**
 * Inspect a specific live R7-Office Desktop editor frame over CDP.
 *
 * Usage:
 *   node scripts/live-cdp-inspect.mjs [--match Documents] [--port 9222]
 *
 * Reports the editor globals, plugin-management surfaces and loaded plugin
 * list, which is what the live bridge test uses to load the bridge plugin
 * programmatically rather than clicking through the UI.
 */

import { waitForCdp, listTargets, CdpSession, delay } from './cdp.mjs'

function argValue(flag, fallback) {
  const i = process.argv.indexOf(flag)
  return i !== -1 && process.argv[i + 1] ? process.argv[i + 1] : fallback
}

const PORT = Number(argValue('--port', process.env.R7_CDP_PORT || '9222'))
const MATCH = argValue('--match', 'Documents')

await waitForCdp(PORT, 20000)

const targets = await listTargets(PORT)
console.log('targets:')
for (const t of targets) {
  console.log(`  - [${t.type}] ${t.title}`)
}

const target = targets.find(t => t.type === 'page' && t.url.includes('/apps/api/documents/'))
  || targets.find(t => t.type === 'page' && t.url.includes('documents/index.html'))
  || targets.find(t => t.type === 'page' && t.title.includes(MATCH))

if (!target) {
  console.error(`no target matching "${MATCH}" found`)
  process.exit(2)
}

console.log(`\nattached to: ${target.title}\n`)

const session = await CdpSession.connect(target.webSocketDebuggerUrl)
await delay(2500)

const probe = await session.evaluate(`(() => {
  const out = {
    href: location.href.slice(0, 160),
    hasAsc: typeof window.Asc !== 'undefined',
    hasDocsAPI: typeof window.DocsAPI !== 'undefined',
    frameCount: window.frames.length,
    frameUrls: []
  };
  for (let i = 0; i < window.frames.length; i++) {
    try { out.frameUrls.push(String(window.frames[i].location.href).slice(0, 120)); }
    catch (e) { out.frameUrls.push('<cross-origin>'); }
  }
  if (out.hasAsc && window.Asc) out.ascKeys = Object.keys(window.Asc).slice(0, 50);
  if (window.Asc && window.Asc.pluginManager) {
    out.pluginManagerKeys = Object.keys(window.Asc.pluginManager).slice(0, 80);
  }
  if (window.Asc && window.Asc.editorConfig) {
    out.editorConfigKeys = Object.keys(window.Asc.editorConfig).slice(0, 50);
  }
  if (window.editorConfig) out.editorConfigGlobal = Object.keys(window.editorConfig).slice(0, 40);
  return out;
})()`)

console.log('probe:', JSON.stringify(probe, null, 2))

// Ask the editor for its installed / connected plugin list when available.
const pluginList = await session.evaluate(`(async () => {
  try {
    if (window.Asc && window.Asc.pluginManager && typeof window.Asc.pluginManager.getPlugins === 'function') {
      const list = await window.Asc.pluginManager.getPlugins();
      return { source: 'getPlugins', count: Array.isArray(list) ? list.length : null,
               names: Array.isArray(list) ? list.map(p => p.name || p.guid).slice(0, 30) : null };
    }
    if (window.Asc && window.Asc.pluginManager && typeof window.Asc.pluginManager.getAllPlugins === 'function') {
      const list = await window.Asc.pluginManager.getAllPlugins();
      return { source: 'getAllPlugins', count: Array.isArray(list) ? list.length : null,
               names: Array.isArray(list) ? list.map(p => p.name || p.guid).slice(0, 30) : null };
    }
  } catch (e) {
    return { error: String(e) };
  }
  return { source: 'none' };
})()`)

console.log('plugin list:', JSON.stringify(pluginList, null, 2))

session.close()
