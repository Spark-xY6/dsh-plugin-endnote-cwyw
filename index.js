import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';

export const name = 'endnote-cwyw-skills';
export const inject = ['skills'];

// These are packaged instructions. Project files never supply a provider locator.
const summaries = new Map([
  ['cite-from-pdfs', 'Use PDFs to support missing citations in a Word manuscript. Match papers to a copy of an EndNote library and deliver both the updated library and editable manuscript.'],
  ['endnote-cwyw', 'Audit, insert or reformat genuine EndNote Cite While You Write fields in a Windows Word document using the official EndNote add-in.'],
]);

function describe(skill, description) {
  const directory = new URL(`./skills/${skill}/`, import.meta.url);
  return {
    name: skill,
    description,
    invocation: { modelInvocable: true, userInvocable: true },
    provider: name,
    source: 'bundled',
    path: fileURLToPath(new URL('SKILL.md', directory)),
    resourceBase: { kind: 'directory', path: fileURLToPath(directory) },
  };
}

export function apply(ctx) {
  ctx.skills.registerProvider(() => ({
    name,
    list() {
      return Array.from(summaries, ([skill, description]) => ({
        ...describe(skill, description),
        rank: 600,
        locator: skill,
      }));
    },
    async get(candidate) {
      const description = summaries.get(candidate.name);
      if (!description || candidate.locator !== candidate.name) {
        throw new Error('Unrecognized bundled skill.');
      }
      const summary = describe(candidate.name, description);
      const markdown = await readFile(summary.path, 'utf8');
      const frontmatter = /^---\r?\n([\s\S]*?)\r?\n---\r?\n/u.exec(markdown);
      if (!frontmatter || !frontmatter[1].split(/\r?\n/u).includes(`name: ${candidate.name}`)) {
        throw new Error(`Invalid frontmatter in ${candidate.name}/SKILL.md`);
      }
      return { ...summary, content: markdown.slice(frontmatter[0].length).trim() };
    },
  }));
}
