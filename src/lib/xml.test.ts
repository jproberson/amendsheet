import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import { readContainer } from './container.js'
import { type XmlEvent, readXml } from './xml.js'

const events = (source: string): XmlEvent[] => [...readXml(source)]

/** Offsets are asserted separately; this keeps shape assertions readable. */
function withoutSpan(event: XmlEvent | undefined) {
  if (event === undefined) return undefined
  const { start: _start, end: _end, ...rest } = event
  return rest
}

const shapes = (source: string) => events(source).map(withoutSpan)

test('reads an element with attributes', () => {
  const [first] = events('<c r="A1" t="s">')

  assert.deepEqual(withoutSpan(first), {
    kind: 'open',
    name: 'c',
    attributes: new Map([
      ['r', 'A1'],
      ['t', 's'],
    ]),
    selfClosing: false,
  })
})

test('marks self closing elements', () => {
  const [first] = events('<f t="shared" si="0"/>')

  assert.equal(first?.kind, 'open')
  assert.equal(first?.kind === 'open' && first.selfClosing, true)
})

test('reads text between elements', () => {
  assert.deepEqual(shapes('<t>hello</t>'), [
    { kind: 'open', name: 't', attributes: new Map(), selfClosing: false },
    { kind: 'text', text: 'hello' },
    { kind: 'close', name: 't' },
  ])
})

test('decodes the predefined entities', () => {
  const [, text] = events('<t>a &amp; b &lt;c&gt; &quot;d&quot; &apos;e&apos;</t>')

  assert.deepEqual(withoutSpan(text), { kind: 'text', text: `a & b <c> "d" 'e'` })
})

test('decodes numeric character references', () => {
  const [, text] = events('<t>&#65;&#x42;&#128512;</t>')

  assert.deepEqual(withoutSpan(text), { kind: 'text', text: 'AB\u{1F600}' })
})

test('decodes entities inside attribute values', () => {
  const [first] = events('<c r="A&amp;1"/>')

  assert.equal(first?.kind === 'open' && first.attributes.get('r'), 'A&1')
})

test('skips the xml declaration, comments and processing instructions', () => {
  const source = '<?xml version="1.0"?><!-- note --><a/>'

  assert.deepEqual(shapes(source), [
    { kind: 'open', name: 'a', attributes: new Map(), selfClosing: true },
  ])
})

test('keeps CDATA content verbatim', () => {
  const [, text] = events('<t><![CDATA[a & <b>]]></t>')

  assert.deepEqual(withoutSpan(text), { kind: 'text', text: 'a & <b>' })
})

test('reports where an unclosed tag started', () => {
  assert.throws(() => events('<a><b'), /offset 3/)
})

test('parses every xml part in the fixtures with balanced elements', async () => {
  const files = (await readdir('fixtures/real')).filter((name) => name.endsWith('.xlsx'))
  assert.ok(files.length > 0, 'fixtures is missing')

  for (const file of files) {
    const container = readContainer(new Uint8Array(await readFile(`fixtures/real/${file}`)))

    for (const [path, bytes] of container.parts) {
      if (!path.endsWith('.xml') && !path.endsWith('.rels')) continue

      const open: string[] = []
      for (const event of readXml(new TextDecoder().decode(bytes))) {
        if (event.kind === 'open' && !event.selfClosing) open.push(event.name)
        if (event.kind === 'close') {
          assert.equal(open.pop(), event.name, `${file} ${path}: mismatched close`)
        }
      }
      assert.deepEqual(open, [], `${file} ${path}: unclosed elements`)
    }
  }
})

test('ignores whitespace before a tag closes', () => {
  const [first] = events('<c r="A1"  />')

  assert.equal(first?.kind === 'open' && first.attributes.get('r'), 'A1')
})

test('ignores an attribute with no name', () => {
  const [first] = events('<c ="x"/>')

  assert.equal(first?.kind === 'open' && first.attributes.size, 0)
})

test('reads text that follows the last element', () => {
  const [, , , tail] = events('<a>x</a>tail')

  assert.deepEqual(withoutSpan(tail), { kind: 'text', text: 'tail' })
})

test('reads a document with no elements', () => {
  assert.deepEqual(shapes('plain'), [{ kind: 'text', text: 'plain' }])
})

test('rejects an attribute with no value', () => {
  assert.throws(() => events('<c r/>'), /has no value/)
})

test('rejects an unquoted attribute value', () => {
  assert.throws(() => events('<c r=A1/>'), /is not quoted/)
})

test('rejects unterminated markup', () => {
  assert.throws(() => events('<!-- open'), /Unterminated comment/)
  assert.throws(() => events('<![CDATA[ open'), /Unterminated CDATA/)
  assert.throws(() => events('<?xml open'), /Unterminated processing instruction/)
  assert.throws(() => events('<!DOCTYPE open'), /Unterminated declaration/)
})

test('reads single quoted attribute values', () => {
  const [first] = events("<c r='A1'/>")

  assert.equal(first?.kind === 'open' && first.attributes.get('r'), 'A1')
})

test('ignores whitespace in a closing tag', () => {
  const [, close] = events('<a></a >')

  assert.deepEqual(withoutSpan(close), { kind: 'close', name: 'a' })
})

test('leaves an unknown entity alone', () => {
  const [, text] = events('<t>&nbsp; &amp;</t>')

  assert.deepEqual(withoutSpan(text), { kind: 'text', text: '&nbsp; &' })
})

test('treats tabs and newlines inside a tag as whitespace', () => {
  const [first] = events('<c\tr="A1"\n\ts="2"\r\n/>')

  assert.equal(first?.kind === 'open' && first.name, 'c')
  assert.equal(first?.kind === 'open' && first.attributes.get('r'), 'A1')
  assert.equal(first?.kind === 'open' && first.attributes.get('s'), '2')
})

test('reports where each event was found', () => {
  const source = '<a><b>hi</b></a>'

  assert.deepEqual(
    events(source).map((event) => [event.start, event.end]),
    [
      [0, 3],
      [3, 6],
      [6, 8],
      [8, 12],
      [12, 16],
    ],
  )
})

test('offsets cover the exact bytes of an element', () => {
  const source = '<row r="1"><c r="A1" t="s"><v>3</v></c></row>'
  const [, open] = events(source)

  assert.equal(source.slice(open?.start, open?.end), '<c r="A1" t="s">')
})

test('offsets skip over comments and declarations', () => {
  const source = '<?xml version="1.0"?><a/>'
  const [first] = events(source)

  assert.equal(source.slice(first?.start, first?.end), '<a/>')
})

test('offsets cover cdata including its wrapper', () => {
  const source = '<t><![CDATA[x]]></t>'
  const [, text] = events(source)

  assert.equal(source.slice(text?.start, text?.end), '<![CDATA[x]]>')
})
