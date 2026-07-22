// Runs the built library inside a real browser, so the browser-support claim is
// checked by execution rather than by a grep for Node APIs. It bundles the ESM
// output with fflate, reads a fixture, edits a cell, writes it back and reads
// the result, all in the page, and asserts no Node globals are in scope.
//
// It drives whatever Chrome is on the machine over the DevTools protocol, using
// Node's built-in WebSocket, so it adds no browser-automation dependency. With
// no Chrome found it exits 0 and says so, rather than failing a machine that
// cannot run it.

import { spawn } from 'node:child_process'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { setTimeout as delay } from 'node:timers/promises'

const CHROME_CANDIDATES = [
  process.env.CHROME_BIN,
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  '/Applications/Chromium.app/Contents/MacOS/Chromium',
  '/usr/bin/google-chrome',
  '/usr/bin/chromium',
  '/usr/bin/chromium-browser',
]

async function findChrome() {
  for (const path of CHROME_CANDIDATES) {
    if (path === undefined) continue
    try {
      await readFile(path)
      return path
    } catch {}
  }
  return undefined
}

function launch(chrome, profile) {
  const child = spawn(
    chrome,
    [
      '--headless=new',
      '--disable-gpu',
      '--no-sandbox',
      '--no-first-run',
      '--no-default-browser-check',
      '--remote-debugging-port=0',
      `--user-data-dir=${profile}`,
      'about:blank',
    ],
    { stdio: ['ignore', 'ignore', 'pipe'] },
  )

  const endpoint = new Promise((resolve, reject) => {
    let buffer = ''
    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk) => {
      buffer += chunk
      const match = buffer.match(/DevTools listening on (ws:\/\/\S+)/)
      if (match) resolve(match[1])
    })
    child.on('exit', (code) => reject(new Error(`Chrome exited early with code ${code}`)))
    delay(15_000).then(() => reject(new Error('Chrome did not report a DevTools endpoint')))
  })

  return { child, endpoint }
}

function connect(url) {
  const socket = new WebSocket(url)
  const pending = new Map()
  let nextId = 1

  socket.addEventListener('message', (event) => {
    const message = JSON.parse(event.data)
    const waiting = pending.get(message.id)
    if (waiting === undefined) return
    pending.delete(message.id)
    if (message.error) waiting.reject(new Error(message.error.message))
    else waiting.resolve(message.result)
  })

  const ready = new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true })
    socket.addEventListener('error', () => reject(new Error('DevTools socket failed')), {
      once: true,
    })
  })

  const send = (method, params = {}, sessionId) =>
    new Promise((resolve, reject) => {
      const id = nextId++
      pending.set(id, { resolve, reject })
      socket.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }))
    })

  return { socket, ready, send }
}

function pageProgram(amendSource, fflateSource, fixtureBase64) {
  return `(async () => {
    const bytesFrom = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    const moduleUrl = (source) =>
      URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));

    const fflateUrl = moduleUrl(${JSON.stringify(fflateSource)});
    const amendSource = ${JSON.stringify(amendSource)}.replace(
      /from\\s*['"]fflate['"]/,
      "from '" + fflateUrl + "'",
    );
    const mod = await import(moduleUrl(amendSource));

    const workbook = mod.readWorkbook(bytesFrom(${JSON.stringify(fixtureBase64)}));
    const sheet = workbook.sheets[0];
    // An empty, unstyled cell, so the value comes back a plain number rather
    // than inheriting a date format the fixture put on the cells it does use.
    sheet.set('Z99', 424242);
    const reopened = mod.readWorkbook(workbook.toBytes());
    const cell = reopened.sheets[0].cell('Z99');

    return {
      wrote: cell && cell.value.kind === 'number' && cell.value.value === 424242,
      kind: cell && cell.value.kind,
      value: cell && cell.value.value,
      noProcess: typeof process === 'undefined',
      noBuffer: typeof Buffer === 'undefined',
      exportNames: Object.keys(mod).sort().join(','),
    };
  })()`
}

async function main() {
  const chrome = await findChrome()
  if (chrome === undefined) {
    console.log('SKIPPED: no Chrome found (set CHROME_BIN to run the browser check)')
    return
  }

  const [amendSource, fflateSource, fixture] = await Promise.all([
    readFile('dist/index.js', 'utf8'),
    readFile('node_modules/fflate/esm/browser.js', 'utf8'),
    readFile('fixtures/quirks/date-epoch-1904.xlsx'),
  ])
  const fixtureBase64 = fixture.toString('base64')

  const profile = await mkdtemp(join(tmpdir(), 'amendsheet-browser-'))
  const { child, endpoint } = launch(chrome, profile)
  let client

  try {
    client = connect(await endpoint)
    await client.ready

    const { targetId } = await client.send('Target.createTarget', { url: 'about:blank' })
    const { sessionId } = await client.send('Target.attachToTarget', { targetId, flatten: true })
    await client.send('Runtime.enable', {}, sessionId)

    const evaluated = await client.send(
      'Runtime.evaluate',
      {
        expression: pageProgram(amendSource, fflateSource, fixtureBase64),
        awaitPromise: true,
        returnByValue: true,
      },
      sessionId,
    )

    if (evaluated.exceptionDetails) {
      const detail = evaluated.exceptionDetails
      throw new Error(`page threw: ${detail.exception?.description ?? detail.text}`)
    }

    const result = evaluated.result.value
    const failures = []
    if (!result.wrote) failures.push(`A1 read back as ${result.kind} ${result.value}, not 424242`)
    if (!result.noProcess) failures.push('process was defined, so this was not a browser context')
    if (!result.noBuffer) failures.push('Buffer was defined, so this was not a browser context')
    if (!result.exportNames.includes('readWorkbook')) {
      failures.push(`the bundle exported ${result.exportNames}`)
    }

    if (failures.length > 0) {
      console.error(`FAILED in ${chrome}:`)
      for (const line of failures) console.error(`  ${line}`)
      process.exitCode = 1
      return
    }

    console.log(`PASSED: read, edited and wrote in the browser (${result.exportNames})`)
  } finally {
    client?.socket.close()
    child.kill()
    await rm(profile, { recursive: true, force: true }).catch(() => {})
  }
}

main().catch((error) => {
  console.error(`browser smoke test errored: ${error.message}`)
  process.exitCode = 1
})
