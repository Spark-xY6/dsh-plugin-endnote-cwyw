import assert from 'node:assert/strict';
import { readFile, stat } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { Context } from '@deepseek-ai/cordis';
import SkillRegistry from '@deepseek-ai/dsh-skill';
import * as bundle from '../index.js';

const root = fileURLToPath(new URL('../', import.meta.url));
const implementations = new Map([
  ['cite-from-pdfs', 'apply_citations.py'],
  ['endnote-cwyw', 'format_cwyw.py'],
]);

test('package entry and locale details resolve from the package directory', async () => {
  const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'));
  const patch = await readFile(join(root, pkg.dsh.bundle.patch), 'utf8');
  assert.match(patch, /^- insert:\s+[-] id: endnote-cwyw-skills\s+name: dsh-plugin-endnote-cwyw\s*$/u);
  assert.equal((await import(pkg.name)).apply, bundle.apply);
  for (const locale of ['en-US', 'zh-CN']) {
    const { meta } = JSON.parse(await readFile(join(root, 'locale', `${locale}.json`), 'utf8'));
    assert.ok(meta.title.trim());
    assert.ok(meta.description.trim());
    const detail = await readFile(join(root, meta.detail), 'utf8');
    assert.ok(/^# /u.test(detail));
  }
  assert.ok((await stat(join(root, pkg.icon))).isFile());
});

test('Cordis mounts both skills, reads packaged bodies and removes them on unload', async () => {
  const ctx = new Context();
  const registry = await ctx.plugin(SkillRegistry);
  try {
    const mounted = await ctx.plugin(bundle);
    try {
      const options = { cwd: dirname(dirname(root)) };
      const snapshot = await ctx.skills.snapshot(options);
      assert.equal(snapshot.complete, true);
      assert.deepEqual(snapshot.skills.map((skill) => skill.name), [...implementations.keys()]);
      for (const [skill, script] of implementations) {
        const loaded = await ctx.skills.get(skill, options);
        assert.equal(loaded.provider, bundle.name);
        assert.equal(loaded.source, 'bundled');
        assert.deepEqual(loaded.invocation, { modelInvocable: true, userInvocable: true });
        assert.equal(loaded.resourceBase.kind, 'directory');
        const base = join(root, 'skills', skill);
        assert.equal(join(loaded.resourceBase.path, 'SKILL.md'), join(base, 'SKILL.md'));
        assert.ok((await stat(join(base, 'scripts', script))).isFile());
        const raw = await readFile(loaded.path, 'utf8');
        assert.ok(loaded.content.startsWith('# '));
        assert.ok(raw.trimEnd().endsWith(loaded.content));
      }
      assert.equal(await ctx.skills.get('unknown-skill', options), undefined);
    } finally {
      await mounted.dispose();
    }
    assert.deepEqual(await ctx.skills.list(), []);
  } finally {
    await registry.dispose();
  }
});
