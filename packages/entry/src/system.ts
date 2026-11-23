/** Safe host-soffice conversion with a private profile and cancellation. */
import { spawn } from 'node:child_process'
import { access, mkdir, rename } from 'node:fs/promises'
import { constants } from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { ConversionError } from './errors.ts'
import type { ConversionSpec } from './operations.ts'
import type { SystemEngine } from './engine.ts'
import type { ResolvedOptions } from './options.ts'

export async function runSystem(engine: SystemEngine, _options: ResolvedOptions, input: string, output: string,
  profile: string, signal: AbortSignal, operation: Pick<ConversionSpec, 'format' | 'recalculate' | 'sheet'>): Promise<void> {
  signal.throwIfAborted()
  await mkdir(profile, { recursive: true, mode: 0o700 })
  const outputDir = dirname(output)
  const args = ['--headless', '--nologo', '--nodefault', '--nofirststartwizard', '--norestore',
    `-env:UserInstallation=file://${profile}`, '--convert-to', operation.format,
    '--outdir', outputDir, input]
  const env: Record<string, string> = {}
  for (const key of ['PATH', 'LANG', 'LC_ALL', 'TZ']) if (process.env[key] !== undefined) env[key] = process.env[key]!
  Object.assign(env, { HOME: profile, USERPROFILE: profile, TMPDIR: profile, TMP: profile, TEMP: profile })
  const child = spawn(engine.executable, args, { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true, detached: true, env })
  let stderr = ''
  let failure: unknown
  const abort = () => {
    failure ??= signal.reason
    try { if (child.pid !== undefined) process.kill(-child.pid, 'SIGKILL') } catch { child.kill('SIGKILL') }
  }
  signal.addEventListener('abort', abort, { once: true })
  child.stderr.setEncoding('utf8')
  child.stderr.on('data', (chunk: string) => { stderr = (stderr + chunk).slice(-16_384) })
  try {
    const code = await new Promise<number | null>((resolve, reject) => {
      child.once('error', error => { failure ??= error })
      child.once('close', (value, killed) => {
        if (killed) failure ??= new Error(`soffice exited by ${killed}`)
        if (failure) reject(failure); else resolve(value)
      })
    })
    signal.throwIfAborted()
    if (code !== 0) throw new ConversionError('failed', `System LibreOffice exited with ${code}: ${stderr}`)
    const produced = join(outputDir, `${basename(input, '.' + input.split('.').pop())}.${operation.format}`)
    try { await access(produced, constants.F_OK) } catch {
      throw new ConversionError('invalid-output', `System LibreOffice exited successfully but produced no ${operation.format} output: ${stderr}`)
    }
    if (produced !== output) await rename(produced, output)
  } finally {
    signal.removeEventListener('abort', abort)
  }
}
