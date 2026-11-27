import { beforeAll } from 'vitest'

beforeAll(() => {
  // Unit tests mock the packaged engines; production auto-selection is tested separately.
  process.env.LIBREOFFICE_KIT_BACKEND = 'wasm'
})
