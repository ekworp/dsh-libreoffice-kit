#!/usr/bin/env node
/** Benchmark one caller-supplied PPTX with cold converters and converter reuse. */
import { createConverter } from '../packages/entry/lib/index.js'
import { mkdir, readFile, rm } from 'node:fs/promises'
import { basename, resolve, join } from 'node:path'

const input = resolve(process.argv[2] ?? '')
const backend = process.argv[3] ?? 'system'
if (!input || !input.endsWith('.pptx')) throw new Error('Usage: target-ppt.mjs /absolute/input.pptx [system|auto|wasm]')
const root = await import('node:fs/promises').then(({ mkdtemp }) => mkdtemp('/tmp/libreoffice-kit-target-'))
const options = { backend, timeoutMs: 120_000, diagnostics: true }
const rows = []
async function one(mode, iteration, converter) {
  const output = join(root, `${mode}-${iteration}.pdf`)
  const started = performance.now()
  try {
    const result = await converter.render({ inputPath: input, outputPath: output })
    const pdf = await readFile(output)
    rows.push({ mode, iteration, ok: true, elapsedMs: performance.now() - started,
      bytes: pdf.length, magic: pdf.subarray(0, 5).toString(), backend: result.backend, timings: result.timings })
  } catch (error) {
    rows.push({ mode, iteration, ok: false, elapsedMs: performance.now() - started,
      code: error?.code, error: error instanceof Error ? error.message : String(error) })
  }
}
for (let i = 0; i < 3; i++) {
  const started = performance.now(); let converter
  try { converter = await createConverter(options); await one('cold', i, converter); rows.at(-1).createMs = performance.now() - started }
  catch (error) { rows.push({ mode: 'cold', iteration: i, ok: false, createMs: performance.now() - started, error: String(error) }) }
  finally { await converter?.dispose() }
}
let converter
try { converter = await createConverter(options); for (let i = 0; i < 3; i++) await one('reuse', i, converter) }
finally { await converter?.dispose(); await rm(root, { recursive: true, force: true }) }
console.log(JSON.stringify({ input: basename(input), backend, rows }, null, 2))
