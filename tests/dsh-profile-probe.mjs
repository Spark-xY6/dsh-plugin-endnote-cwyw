import assert from 'node:assert/strict';
import { stat } from 'node:fs/promises';
import { join } from 'node:path';

export const inject = ['appReady', 'appExit', 'skills'];

export function apply(ctx) {
  ctx.appReady.onReady(async () => {
    try {
      const expected = {
        'cite-from-pdfs': 'apply_citations.py',
        'endnote-cwyw': 'format_cwyw.py',
      };
      const catalog = await ctx.skills.snapshot();
      assert.equal(catalog.complete, true, 'DSH returned an incomplete skill catalog');
      for (const [name, script] of Object.entries(expected)) {
        assert.ok(catalog.skills.some((item) => item.name === name));
        const skill = await ctx.skills.get(name);
        assert.equal(skill.provider, 'endnote-cwyw-skills');
        assert.equal(skill.resourceBase.kind, 'directory');
        assert.ok(skill.content.startsWith('# '));
        assert.ok((await stat(join(skill.resourceBase.path, 'scripts', script))).isFile());
      }
      console.log('DSH_PROFILE_PASS cite-from-pdfs endnote-cwyw');
      ctx.appExit(0);
    } catch (error) {
      console.error(error.message);
      ctx.appExit(1);
    }
  });
}
