import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { runInNewContext } from 'node:vm'
import { describe, expect, it } from 'vitest'

/* forms.js is split into pure helpers (field visibility, answer building,
   validation) and thin DOM glue. These tests run the pure helpers in a VM
   context with no DOM, mirroring the opencode.test.ts harness style. */

function harness(): Record<string, unknown> {
  const context: Record<string, unknown> = {}
  const source = readFileSync(join(__dirname, 'forms.js'), 'utf8')
  runInNewContext(source, context)
  return context
}

const F = () =>
  harness().OpenCodeForms as {
    visibleFields: (fields: unknown[], answers: Record<string, unknown>) => unknown[]
    buildAnswer: (
      fields: unknown[],
      state: { values: Record<string, unknown>; custom: Record<string, string> },
    ) => Record<string, unknown>
    validate: (
      fields: unknown[],
      state: { values: Record<string, unknown>; custom: Record<string, string> },
    ) => string | null
  }

describe('visibleFields', () => {
  it('drops hidden fields', () => {
    const fields = [
      { key: 'a', type: 'string', hidden: true },
      { key: 'b', type: 'string' },
    ]
    expect(
      F()
        .visibleFields(fields, {})
        .map((f) => (f as { key: string }).key),
    ).toEqual(['b'])
  })

  it('honours when conditions (eq / neq) against current answers', () => {
    const fields = [
      { key: 'mode', type: 'string', options: [{ value: 'a' }, { value: 'b' }] },
      { key: 'dep', type: 'string', when: [{ key: 'mode', op: 'eq', value: 'a' }] },
      { key: 'other', type: 'string', when: [{ key: 'mode', op: 'neq', value: 'a' }] },
    ]
    const forms = F()
    expect(forms.visibleFields(fields, { mode: 'a' }).map((f) => (f as { key: string }).key)).toEqual(['mode', 'dep'])
    expect(forms.visibleFields(fields, { mode: 'b' }).map((f) => (f as { key: string }).key)).toEqual(['mode', 'other'])
    // A field with `when` whose dependency is unanswered stays hidden.
    expect(forms.visibleFields(fields, {}).map((f) => (f as { key: string }).key)).toEqual(['mode'])
  })
})

describe('buildAnswer', () => {
  it('maps each field type to its Form.Value shape', () => {
    const fields = [
      { key: 'pick', type: 'string', options: [{ value: 'x' }, { value: 'y' }] },
      { key: 'multi', type: 'multiselect', options: [{ value: '1' }, { value: '2' }] },
      { key: 'flag', type: 'boolean', default: true },
      { key: 'count', type: 'integer' },
      { key: 'ratio', type: 'number' },
      { key: 'note', type: 'string' },
    ]
    const state = {
      values: { pick: 'x', multi: ['1', '2'], flag: true, count: '3', ratio: '0.5', note: 'hi' },
      custom: {},
    }
    expect(F().buildAnswer(fields, state)).toEqual({
      pick: 'x',
      multi: ['1', '2'],
      flag: true,
      count: 3,
      ratio: 0.5,
      note: 'hi',
    })
  })

  it('appends the custom "other" text to a multiselect answer', () => {
    const fields = [{ key: 'multi', type: 'multiselect', options: [{ value: '1' }], custom: true }]
    const state = { values: { multi: ['1'] }, custom: { multi: 'custom answer' } }
    expect(F().buildAnswer(fields, state)).toEqual({ multi: ['1', 'custom answer'] })
    // Custom text alone (nothing checked) still answers with one entry.
    const only = { values: { multi: [] }, custom: { multi: 'solo' } }
    expect(F().buildAnswer(fields, only)).toEqual({ multi: ['solo'] })
  })

  it('omits unanswered optional fields', () => {
    const fields = [
      { key: 'req', type: 'string' },
      { key: 'opt', type: 'string' },
    ]
    const state = { values: { req: 'yes' }, custom: {} }
    expect(F().buildAnswer(fields, state)).toEqual({ req: 'yes' })
  })
})

describe('validate', () => {
  it('flags an empty required string', () => {
    const fields = [{ key: 'q', type: 'string', required: true }]
    expect(F().validate(fields, { values: {}, custom: {} })).toBeTruthy()
    expect(F().validate(fields, { values: { q: 'x' }, custom: {} })).toBeNull()
  })

  it('enforces multiselect minItems / maxItems', () => {
    const fields = [
      {
        key: 'm',
        type: 'multiselect',
        options: [{ value: '1' }, { value: '2' }, { value: '3' }],
        required: true,
        minItems: 1,
        maxItems: 2,
      },
    ]
    const forms = F()
    expect(forms.validate(fields, { values: { m: [] }, custom: {} })).toBeTruthy()
    expect(forms.validate(fields, { values: { m: ['1'] }, custom: {} })).toBeNull()
    expect(forms.validate(fields, { values: { m: ['1', '2', '3'] }, custom: {} })).toBeTruthy()
  })

  it('ignores validation for fields hidden by `when`', () => {
    const fields = [
      { key: 'mode', type: 'string', options: [{ value: 'a' }, { value: 'b' }] },
      { key: 'dep', type: 'string', required: true, when: [{ key: 'mode', op: 'eq', value: 'a' }] },
    ]
    // mode=b hides `dep`, so its missing answer must not block submission.
    expect(F().validate(fields, { values: { mode: 'b' }, custom: {} })).toBeNull()
  })
})
