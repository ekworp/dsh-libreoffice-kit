/** Convert caller-authorized disk documents with installed native or Node WASM engines. */
import { constants } from 'node:fs'
import { open, mkdtemp, mkdir, rm, unlink, writeFile } from 'node:fs/promises'
import type { FileHandle } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Worker } from 'node:worker_threads'
import { resolveImageRender } from './image-operations.ts'
import { inspectDocument } from './document.ts'
import { createFontLoader, preloadFonts } from './font-loader.ts'
import type { ImageRenderSpec, RenderImagesRequest, RenderImagesResult } from './image-operations.ts'
export { IMAGE_FORMATS } from './image-operations.ts'
export type { RenderImagesRequest, RenderImagesResult, RenderedImage } from './image-operations.ts'
import { runNative } from './native.ts'
import { renderImagesWithNative } from './native-image-renderer.ts'
import { resolveOptions } from './options.ts'
import { resolveEngine } from './engine.ts'
import { ConversionError, failureCode } from './errors.ts'
import { pdfPageCount, resolveConversion, validateOutput } from './operations.ts'
import type { ConversionRequest, ConversionSpec } from './operations.ts'
export { CONVERSION_FORMATS } from './operations.ts'
export type { ConversionRequest } from './operations.ts'
export { discoverRuntime } from './runtime.ts'
export type { RuntimeInfo } from './runtime.ts'
import { resolveSystemEngine } from './engine.ts'
import type { Engine, EnginePackageResolver, NativeEngine } from './engine.ts'
export type { EnginePackageResolver } from './engine.ts'
import { runSystem } from './system.ts'
import { FontSnapshotState } from './font-snapshot-state.ts'
import { validateFontSnapshot } from './font-snapshot-validation.ts'
import type { FontSubstitution } from './font-loader.ts'
import type { FontResolutionCacheEntry } from './font-loader.ts'
import type { WorkerRequest } from './worker.ts'
export { ConversionError } from './errors.ts'
export type { ConversionErrorCode } from './errors.ts'
export { ENGINE_VERSION, ENGINE_VERSIONS } from './engine.ts'

/** Host limits used by both native and Node WASM conversions. All byte limits are positive safe integers. */
export interface ConverterOptions {
  /** Engine package manifest resolver. Defaults to Node resolution from this package. */
  resolveEnginePackage?: EnginePackageResolver
  /** Engine selection. `auto` preserves platform package resolution; `system` explicitly uses host soffice on Linux. */
  backend?: 'auto' | 'wasm' | 'native' | 'system'
  /** Include per-stage performance timings in each render result. */
  diagnostics?: boolean
  /** Keep an idle WASM worker alive for this many milliseconds. Default: 30000. */
  wasmWorkerIdleMs?: number
  /** User-local font metadata cache directory. Defaults to the platform cache directory; false disables disk reuse. */
  fontMetadataCacheDirectory?: string | false
  /** Maximum persisted and reusable per-file metadata bytes. Default: 32 MiB. */
  maxFontMetadataCacheBytes?: number
  /** Deadline after a queued conversion starts, including font work and output. Default: 120000 ms. */
  timeoutMs?: number
  /** Maximum source bytes. Default: 64 MiB. */
  maxInputBytes?: number
  /** Maximum generated output bytes. Default: 128 MiB. */
  maxOutputBytes?: number
  /** PDF export image resolution in DPI. Default: 144. */
  maxImageResolution?: number
  /** Maximum ZIP entries. Default: 20000. */
  maxArchiveEntries?: number
  /** Maximum declared uncompressed archive bytes. Default: 512 MiB. */
  maxUncompressedBytes?: number
  /** Font roots; defaults to conventional platform and user directories. */
  fontDirectories?: string[]
  /** Ordered family groups replace the defaults. Exact installed families precede alternatives; the full indexed
   * catalog remains available for missing glyphs. Native LibreOffice also resolves installed metric-compatible
   * fonts before configured choices. */
  fontFallbacks?: string[][]
  /** Families imported before loading a document. Defaults to none. */
  initialFontFamilies?: string[]
  /** Maximum visited physical font files. Default: 20000. */
  maxFontFiles?: number
  /** Files exceeding this limit are omitted. Default: 256 MiB. */
  maxFontFileBytes?: number
  /** Complete original font bytes imported per conversion. Default: 512 MiB. Native platform fonts remain OS-managed. */
  maxLoadedFontBytes?: number
  /** Maximum font match records retained by each compatible process cache or explicit factory. Default: 4096. */
  maxFontResolutionEntries?: number
}

/** Result of one conversion. */
export interface RenderTimings {
  readonly engineDiscoveryMs?: number
  readonly inputReadMs?: number
  readonly fontScanMs?: number
  readonly workerStartMs?: number
  readonly wasmStartupMs?: number
  readonly conversionMs?: number
  readonly pdfOutputReadMs?: number
  readonly outputWriteMs?: number
}

export interface RenderResult {
  readonly backend: 'native' | 'wasm' | 'system'
  /** Number of pages in PDF output when the generated PDF exposes page dictionaries. */
  readonly pageCount?: number
  readonly timings?: RenderTimings
  /** Declared OOXML families no installed font provides; binary formats return an empty list. */
  readonly missingFonts: string[]
}

/** A converter serializes document operations; cancellation and disposal await process/worker exit. */
export interface Converter {
  readonly backend: 'native' | 'wasm' | 'system'
  /**
   * Convert a private, caller-authorized regular Office file to a fresh exclusive PDF path.
   * The caller owns both directories and must prevent concurrent path changes.
   * Rejects existing output paths; removes a newly created output on failure or cancellation.
   * @param request - Absolute paths. Input suffixes are listed in CONVERSION_FORMATS.
   * @param signal - Optional cancellation, including while queued.
   * @returns Engine choice and missing declared font families.
   */
  render(request: { inputPath: string; outputPath: string }, signal?: AbortSignal): Promise<RenderResult>
  /** Render Office directly to PNG, or PDF through PDFium, in one fresh output directory.
   * Worksheets use sheet/A1 coordinates; page selections are one-based for documents and slides.
   * Failure/cancellation removes the complete batch. The manifest describes one saved input snapshot.
   */
  renderImages(request: RenderImagesRequest, signal?: AbortSignal): Promise<RenderImagesResult>
  /** Convert to the output extension; CSV requires a sheet for multi-sheet inputs. */
  convert(request: ConversionRequest, signal?: AbortSignal): Promise<RenderResult>
  /**
   * Recalculate the entire workbook synchronously and save formulas with refreshed cached results.
   * This does not validate formulas or business data.
   * @param request - XLS, XLSX, or ODS input and a distinct fresh XLSX or ODS output.
   * @param signal - Cancellation, including while queued.
   * @returns Engine choice and missing declared font families.
   */
  recalculate(request: Omit<ConversionRequest, 'sheet'>, signal?: AbortSignal): Promise<RenderResult>
  /** Abort queued and active work, await exit and cleanup, and permanently reject further renders. */
  dispose(): Promise<void>
}

/** Factory-scoped engine and font metadata shared by concurrent converters. */
export interface ConverterFactory {
  readonly backend: 'native' | 'wasm' | 'system'
  /** Create one serial converter sharing this factory's bounded font cache. */
  create(): Promise<Converter>
  /** Abort and join every converter created by this factory. */
  dispose(): Promise<void>
}

/** Worker result for a WASM conversion, which returns owned output bytes. */
interface WasmWorkerResult {
  readonly kind?: undefined
  readonly ok: true
  readonly output: Uint8Array
  readonly missingFonts: string[]
  readonly timings?: RenderTimings
}

interface ImagesWorkerResult { readonly kind?: undefined; readonly ok: true; readonly images: RenderImagesResult }

/** Worker result for a native conversion, which returns the font imports the helper loads. */
interface NativeWorkerResult {
  readonly kind?: undefined
  readonly ok: true
  readonly fonts: string[]
  readonly substitutions: FontSubstitution[]
  readonly missingFonts: string[]
  readonly timings?: RenderTimings
}

/** One message the conversion worker sends before it settles. */
type WorkerMessage =
  | { readonly kind: 'font-cache'; readonly entries: FontResolutionCacheEntry[] }
  | { readonly kind: 'phase'; readonly phase: string }
  | WasmWorkerResult
  | ImagesWorkerResult
  | NativeWorkerResult
  | { readonly kind?: undefined; readonly ok: false; readonly code?: string; readonly error: string }

/** File identity that must match across a bounded read. */
const CHANGED_KEYS = ['size', 'mtimeMs', 'ctimeMs'] as const

async function readBounded(path: string, limit: number, signal: AbortSignal, role: 'input' | 'output'): Promise<Buffer> {
  const file = await open(path, constants.O_RDONLY | constants.O_NONBLOCK | constants.O_NOFOLLOW)
  try {
    const before = await file.stat()
    const invalid = role === 'input' ? 'invalid-document' : 'invalid-output'
    if (!before.isFile() || (role === 'input' && before.size < 1)) throw new ConversionError(invalid, 'Document must be a nonempty regular file.')
    if (before.size > limit) throw new ConversionError(`${role}-too-large`, `Document exceeds its ${role} byte limit.`)
    const bytes = Buffer.alloc(before.size)
    for (let offset = 0; offset < bytes.length;) {
      signal.throwIfAborted()
      const { bytesRead } = await file.read(bytes, offset, bytes.length - offset, offset)
      if (!bytesRead) throw new ConversionError(invalid, 'Document was truncated while reading.')
      offset += bytesRead
    }
    const after = await file.stat()
    if (CHANGED_KEYS.some(key => before[key] !== after[key])) throw new ConversionError(invalid, 'Document changed while reading.')
    signal.throwIfAborted()
    return bytes
  } finally { await file.close() }
}

interface WorkerSlot {
  worker: Worker | undefined
  idleTimer: ReturnType<typeof setTimeout> | undefined
}

async function stopWorker(slot: WorkerSlot): Promise<void> {
  const worker = slot.worker
  slot.worker = undefined
  if (slot.idleTimer !== undefined) clearTimeout(slot.idleTimer)
  slot.idleTimer = undefined
  await worker?.terminate()
}

async function runWorker(data: WorkerRequest, signal: AbortSignal,
  onFontCache: (entries: FontResolutionCacheEntry[]) => void, slot?: WorkerSlot, idleMs = 30_000):
Promise<WasmWorkerResult | NativeWorkerResult | ImagesWorkerResult> {
  signal.throwIfAborted()
  if (slot?.idleTimer !== undefined) { clearTimeout(slot.idleTimer); slot.idleTimer = undefined }
  const worker = slot?.worker ?? new Worker(new URL('./worker.js', import.meta.url), {
    workerData: { ...data, ...(slot !== undefined ? { persistent: true } : {}) }, name: 'libreoffice-conversion', execArgv: [], stdout: true, stderr: true,
  })
  const persistent = slot !== undefined && typeof (worker as unknown as { postMessage?: unknown }).postMessage === 'function'
  if (persistent && slot?.worker === worker) worker.postMessage(data)
  if (persistent) slot!.worker = worker
  worker.stdout.resume()
  worker.stderr.resume()
  const settled = Promise.withResolvers<WasmWorkerResult | NativeWorkerResult | ImagesWorkerResult>()
  const abort = (): void => { settled.reject(signal.reason) }
  const error = (cause: unknown): void => { if (slot?.worker === worker) slot.worker = undefined; settled.reject(cause) }
  const exit = (code: number): void => { error(new ConversionError('failed', `LibreOffice worker exited before returning a result (${code}).`)) }
  const message = (value: WorkerMessage): void => {
    if (value.kind === 'font-cache' || value.kind === 'phase') {
      if (value.kind === 'font-cache') onFontCache(value.entries)
    } else if (value.ok) settled.resolve(value)
    else settled.reject(new ConversionError(failureCode(value), value.error))
  }
  signal.addEventListener('abort', abort, { once: true })
  worker.once('error', error)
  worker.once('exit', exit)
  worker.on('message', message)
  let keep = false
  try {
    const result = await settled.promise
    keep = persistent
    return result
  } finally {
    signal.removeEventListener('abort', abort)
    worker.removeListener('error', error)
    worker.removeListener('exit', exit)
    worker.removeListener('message', message)
    if (!keep) await worker.terminate()
    else {
      if (slot!.idleTimer !== undefined) clearTimeout(slot!.idleTimer)
      slot!.idleTimer = setTimeout(() => { void stopWorker(slot!) }, idleMs)
    }
  }
}

/**
 * Resolve the installed engine and deployment limits without starting LibreOffice.
 * @param options - Conversion, font, and output limits.
 * @returns A serial converter; dispose it after use.
 * @throws TypeError for invalid options, or an unavailable ConversionError for invalid installed engines.
 * An absent native engine or a known unsupported glibc version selects WASM.
 */
interface SharedConverterState {
  readonly snapshots: FontSnapshotState
  generation: string | undefined
  readonly fontCache: Map<string, FontResolutionCacheEntry>
}

const DEFAULT_SHARED_STATE_LIMIT = 16
const defaultSharedStates = new Map<string, SharedConverterState>()

function sharedStateKey(options: ReturnType<typeof resolveOptions>, engine: Engine): string {
  return JSON.stringify({ engine, fontDirectories: options.fontDirectories, includeOfficeFonts: options.includeOfficeFonts,
    fontFallbacks: options.fontFallbacks, initialFontFamilies: options.initialFontFamilies,
    maxFontFiles: options.maxFontFiles, maxFontFileBytes: options.maxFontFileBytes,
    maxLoadedFontBytes: options.maxLoadedFontBytes, maxFontResolutionEntries: options.maxFontResolutionEntries,
    fontMetadataCacheDirectory: options.fontMetadataCacheDirectory, maxFontMetadataCacheBytes: options.maxFontMetadataCacheBytes })
}

function defaultSharedState(options: ReturnType<typeof resolveOptions>, engine: Engine): SharedConverterState {
  const key = sharedStateKey(options, engine)
  const existing = defaultSharedStates.get(key)
  if (existing !== undefined) {
    defaultSharedStates.delete(key)
    defaultSharedStates.set(key, existing)
    return existing
  }
  const state: SharedConverterState = { snapshots: new FontSnapshotState(), generation: undefined, fontCache: new Map() }
  defaultSharedStates.set(key, state)
  while (defaultSharedStates.size > DEFAULT_SHARED_STATE_LIMIT)
    defaultSharedStates.delete(defaultSharedStates.keys().next().value as string)
  return state
}

async function resolveConverter(options?: ConverterOptions): Promise<{
  readonly options: ReturnType<typeof resolveOptions>
  readonly engine: Engine
  readonly engineDiscoveryMs: number
}> {
  const started = performance.now()
  const resolvedOptions = resolveOptions(options)
  let engine: Engine
  try {
    if (resolvedOptions.backend === 'system') engine = await resolveSystemEngine()
    else if (resolvedOptions.backend === 'auto') {
      try { engine = await resolveSystemEngine() }
      catch { engine = await resolveEngine(options?.resolveEnginePackage) }
    } else engine = await resolveEngine(options?.resolveEnginePackage)
    if (resolvedOptions.backend !== 'auto' && resolvedOptions.backend !== engine.backend)
      throw new Error(`Requested ${resolvedOptions.backend} backend is unavailable on this host.`)
  } catch (cause) { throw new ConversionError('unavailable', (cause as Error).message, { cause }) }
  return { options: resolvedOptions, engine, engineDiscoveryMs: performance.now() - started }
}

function createConverterWith(resolvedOptions: ReturnType<typeof resolveOptions>, engine: Engine,
  engineDiscoveryMs = 0,
  shared: SharedConverterState, onDispose?: (converter: Converter) => void): Converter {
  const lifetime = new AbortController()
  const active = new Set<Promise<RenderResult | RenderImagesResult>>()
  const waiting = new Set<() => void>()
  let running = false
  let disposal: Promise<void> | undefined
  const workerSlot: WorkerSlot = { worker: undefined, idleTimer: undefined }
  const mergeCache = (generation: string, entries: FontResolutionCacheEntry[]): void => {
    if (shared.generation !== generation) return
    for (const entry of entries) {
      const prior = shared.fontCache.get(entry.key)
      const faces = new Map((prior?.faces ?? []).map(face => [`${face.path}\0${face.faceIndex}`, face]))
      for (const face of entry.faces) faces.set(`${face.path}\0${face.faceIndex}`, face)
      const merged: FontResolutionCacheEntry = { ...entry,
        codePoints: [...new Set([...(prior?.codePoints ?? []), ...entry.codePoints])].sort((left, right) => left - right),
        emptyRequest: prior?.emptyRequest === true || entry.emptyRequest,
        faces: [...faces.values()],
        ...((prior?.missingFamily ?? entry.missingFamily) === undefined ? {}
          : { missingFamily: prior?.missingFamily ?? entry.missingFamily }) }
      shared.fontCache.delete(entry.key)
      shared.fontCache.set(entry.key, merged)
    }
    while (shared.fontCache.size > resolvedOptions.maxFontResolutionEntries)
      shared.fontCache.delete(shared.fontCache.keys().next().value as string)
  }
  async function snapshot(signal: AbortSignal) {
    const current = await shared.snapshots.acquire(resolvedOptions, signal)
    if (shared.generation !== current.generation) {
      shared.generation = current.generation
      shared.fontCache.clear()
    }
    return current
  }
  async function acquire(signal: AbortSignal): Promise<void> {
    while (running) {
      signal.throwIfAborted()
      await new Promise<void>((resolve, reject) => {
        const cleanup = () => { waiting.delete(wake); signal.removeEventListener('abort', abort) }
        const wake = () => { cleanup(); resolve() }
        const abort = () => { cleanup(); reject(signal.reason as Error) }
        waiting.add(wake)
        signal.addEventListener('abort', abort, { once: true })
      })
    }
    signal.throwIfAborted()
    running = true
  }
  async function execute(request: ConversionSpec, signal: AbortSignal): Promise<RenderResult> {
    const { extension, format } = request
    await acquire(signal)
    const deadline = new AbortController()
    const stopped = AbortSignal.any([signal, deadline.signal])
    const timer = setTimeout(() => { deadline.abort(new ConversionError('timeout', 'LibreOffice conversion timed out.')) }, resolvedOptions.timeoutMs)
    let scratch: string | undefined
    let output: FileHandle | undefined
    let succeeded = false
    let failure: unknown
    const timings: { -readonly [K in keyof RenderTimings]?: number } = resolvedOptions.diagnostics ? { engineDiscoveryMs } : {}
    try {
      stopped.throwIfAborted()
      output = await open(request.outputPath, 'wx', 0o600)
      const inputStarted = performance.now()
      const sourceBytes = await readBounded(request.inputPath, resolvedOptions.maxInputBytes, stopped, 'input')
      timings.inputReadMs = performance.now() - inputStarted
      scratch = await mkdtemp(join(tmpdir(), 'libreoffice-kit-'))
      const inputPath = join(scratch, `document.${extension}`)
      const input = await open(inputPath, 'wx', 0o600)
      try { await input.writeFile(sourceBytes) } finally { await input.close() }
      const fontScanStarted = performance.now()
      const fonts = await snapshot(stopped)
      timings.fontScanMs = performance.now() - fontScanStarted
      let bytes: Uint8Array
      let missingFonts: string[] = []
      if (engine.backend === 'system') {
        const preflightStarted = performance.now()
        const document = inspectDocument(sourceBytes, extension, resolvedOptions)
        const loader = createFontLoader(resolvedOptions, document, () => '', fonts.faces, [...shared.fontCache.values()], false)
        preloadFonts(loader, resolvedOptions, document, true)
        missingFonts = loader.missingFonts
        mergeCache(fonts.generation, loader.cacheEntries)
        timings.fontScanMs = (timings.fontScanMs ?? 0) + performance.now() - preflightStarted
        const profile = join(scratch, 'profile')
        const path = join(scratch, `output.${format}`)
        const conversionStarted = performance.now()
        await runSystem(engine, resolvedOptions, inputPath, path, profile, stopped, request)
        timings.conversionMs = performance.now() - conversionStarted
        const outputReadStarted = performance.now()
        bytes = await readBounded(path, resolvedOptions.maxOutputBytes, stopped, 'output')
        timings.pdfOutputReadMs = performance.now() - outputReadStarted
        timings.workerStartMs = 0
      } else {
        const workerStarted = performance.now()
        const result = await runWorker({ inputPath, extension, operation: request, options: resolvedOptions, engine, scratch,
          fontFaces: fonts.faces, fontCache: [...shared.fontCache.values()] }, stopped,
        entries => mergeCache(fonts.generation, entries), workerSlot, resolvedOptions.wasmWorkerIdleMs)
        timings.workerStartMs = performance.now() - workerStarted
        if ('timings' in result && result.timings !== undefined) Object.assign(timings, result.timings)
        if ('images' in result) throw new ConversionError('failed', 'Conversion worker returned an image batch for an export request.')
        missingFonts = result.missingFonts
        if ('output' in result) bytes = result.output
        else {
          const native = engine as NativeEngine
          const profile = join(scratch, 'profile')
          await mkdir(profile, { mode: 0o700 })
          const path = join(scratch, `output.${format}`)
          await runNative(native, resolvedOptions, inputPath, path, profile, result.fonts, result.substitutions, stopped, request)
          bytes = await readBounded(path, resolvedOptions.maxOutputBytes, stopped, 'output')
        }
      }
      validateOutput(bytes, format)
      await validateFontSnapshot(fonts)
      stopped.throwIfAborted()
      const outputWriteStarted = performance.now()
      await output.writeFile(bytes)
      timings.outputWriteMs = performance.now() - outputWriteStarted
      stopped.throwIfAborted()
      succeeded = true
      const pageCount = format === 'pdf' ? pdfPageCount(bytes) : undefined
      return { backend: engine.backend, ...(pageCount === undefined ? {} : { pageCount }), missingFonts,
        ...(resolvedOptions.diagnostics ? { timings } : {}) }
    } catch (error) {
      failure = error
      throw error
    } finally {
      clearTimeout(timer)
      try {
        const errors: unknown[] = []
        try { await output?.close() } catch (error) { errors.push(error) }
        const cleanup = await Promise.allSettled([
          ...(output && !succeeded ? [unlink(request.outputPath)] : []),
          ...(scratch ? [rm(scratch, { recursive: true, force: true })] : []),
        ])
        for (const result of cleanup) {
          if (result.status === 'rejected') errors.push(result.reason)
        }
        if (errors.length) throw new AggregateError(failure ? [failure, ...errors] : errors, 'LibreOffice conversion cleanup failed.')
      } finally {
        running = false
        for (const wake of [...waiting]) wake()
      }
    }
  }
  async function executeImages(request: ImageRenderSpec, signal: AbortSignal): Promise<RenderImagesResult> {
    await acquire(signal)
    const deadline = new AbortController()
    const stopped = AbortSignal.any([signal, deadline.signal])
    const timer = setTimeout(() => deadline.abort(new ConversionError('timeout', 'LibreOffice image rendering timed out.')), resolvedOptions.timeoutMs)
    let scratch: string | undefined
    let owned = false
    let succeeded = false
    try {
      stopped.throwIfAborted()
      await mkdir(request.outputDir, { mode: 0o700 })
      owned = true
      const source = await readBounded(request.inputPath, resolvedOptions.maxInputBytes, stopped, 'input')
      scratch = await mkdtemp(join(tmpdir(), 'libreoffice-kit-images-'))
      const inputPath = join(scratch, `document.${request.extension}`)
      await writeFile(inputPath, source, { flag: 'wx', mode: 0o600 })
      const fonts = await snapshot(stopped)
      const result = await runWorker({ inputPath, extension: request.extension, operation: request, options: resolvedOptions,
        engine, scratch, fontFaces: fonts.faces, fontCache: [...shared.fontCache.values()] }, stopped,
      entries => mergeCache(fonts.generation, entries))
      let images: RenderImagesResult
      if ('images' in result) images = result.images
      else if ('fonts' in result) {
        const profile = join(scratch, 'profile')
        images = await renderImagesWithNative({ engine: engine as NativeEngine, options: resolvedOptions, inputPath,
          source, scratch, profile, fonts: result.fonts, substitutions: result.substitutions,
          missingFonts: result.missingFonts, operation: request, signal: stopped })
      } else throw new ConversionError('failed', 'Image worker did not return its manifest.')
      await validateFontSnapshot(fonts)
      stopped.throwIfAborted()
      await writeFile(join(request.outputDir, 'manifest.json'), `${JSON.stringify(images, null, 2)}\n`, { flag: 'wx', mode: 0o600 })
      stopped.throwIfAborted()
      succeeded = true
      return images
    } finally {
      clearTimeout(timer)
      try {
        const cleanup = await Promise.allSettled([
          ...(scratch ? [rm(scratch, { recursive: true, force: true })] : []),
          ...(owned && !succeeded ? [rm(request.outputDir, { recursive: true, force: true })] : []),
        ])
        const failures = cleanup.filter(result => result.status === 'rejected')
        if (failures.length) throw new AggregateError(failures.map(result => result.reason), 'Image batch cleanup failed.')
      } finally { running = false; for (const wake of [...waiting]) wake() }
    }
  }
  async function submit(request: ConversionRequest, operation: 'render' | 'convert' | 'recalculate', signal?: AbortSignal): Promise<RenderResult> {
    const task = execute(resolveConversion(request, operation),
      signal ? AbortSignal.any([signal, lifetime.signal]) : lifetime.signal)
    active.add(task)
    void task.finally(() => { active.delete(task) }).catch((_error: unknown) => { /* The caller owns task rejection. */ })
    return task
  }
  const converter: Converter = {
    backend: engine.backend,
    render: (request, signal) => submit(request, 'render', signal),
    async renderImages(request, signal) {
      const task = executeImages(resolveImageRender(request), signal ? AbortSignal.any([signal, lifetime.signal]) : lifetime.signal)
      active.add(task)
      void task.finally(() => active.delete(task)).catch(() => {})
      return task
    },
    convert: (request, signal) => submit(request, 'convert', signal),
    recalculate: (request, signal) => submit(request, 'recalculate', signal),
    dispose() {
      disposal ??= (async () => {
        lifetime.abort(new Error('LibreOffice converter is disposed.'))
        await Promise.allSettled(active)
        await stopWorker(workerSlot)
        onDispose?.(converter)
      })()
      return disposal
    },
  }
  return converter
}

/** Resolve one engine and share its bounded font state across many converters. */
export async function createConverterFactory(options?: ConverterOptions): Promise<ConverterFactory> {
  const { options: resolvedOptions, engine, engineDiscoveryMs } = await resolveConverter(options)
  const shared: SharedConverterState = { snapshots: new FontSnapshotState(), generation: undefined, fontCache: new Map() }
  const converters = new Set<Converter>()
  let disposal: Promise<void> | undefined
  const factory: ConverterFactory = {
    backend: engine.backend,
    async create() {
      if (disposal) throw new ConversionError('unavailable', 'LibreOffice converter factory is disposed.')
      const converter = createConverterWith(resolvedOptions, engine, engineDiscoveryMs, shared, value => converters.delete(value))
      converters.add(converter)
      return converter
    },
    dispose() {
      disposal ??= (async () => {
        const active = [...converters]
        converters.clear()
        const results = await Promise.allSettled(active.map(converter => converter.dispose()))
        shared.fontCache.clear()
        shared.snapshots.clear()
        shared.generation = undefined
        const failures = results.filter(result => result.status === 'rejected')
        if (failures.length) throw new AggregateError(failures.map(result => result.reason), 'LibreOffice converter factory disposal failed.')
      })()
      return disposal
    },
  }
  return factory
}

/** Create one converter with process-shared, configuration-isolated bounded font metadata. */
export async function createConverter(options?: ConverterOptions): Promise<Converter> {
  const { options: resolvedOptions, engine, engineDiscoveryMs } = await resolveConverter(options)
  return createConverterWith(resolvedOptions, engine, engineDiscoveryMs, defaultSharedState(resolvedOptions, engine))
}
