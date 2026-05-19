# Update Guide

## When to update the skill pack
Update the pack when:
- Max's role or responsibilities change.
- A recurring work process becomes stable.
- A new business project starts.
- A content format becomes repeatable.
- Max changes tone, language, or output preferences.
- A skill produces outputs that are too generic or no longer useful.

## How to update memory files
- `memory/max-profile.md`: durable professional profile and technical focus.
- `memory/max-preferences.md`: response formats, document preferences, script expectations.
- `memory/max-projects.md`: active project categories and reusable business/content themes.
- `memory/max-tone-style.md`: English and Chinese tone guidance.

## How to update skills
1. Identify the workflow.
2. Add trigger conditions under "When to use this skill".
3. Add or refine decision logic.
4. Add one realistic example in `examples.md`.
5. Add one reusable template in `templates.md` if the format will be reused.
6. Check privacy boundaries.

## Versioning suggestion
Use semantic versioning in release notes if publishing as a GitHub repository:
- Patch: wording or template fixes
- Minor: new examples, prompts, or templates
- Major: new skill architecture or major behavior changes

## Quality checklist after updates
- Does the skill produce a practical output?
- Is the output copy-paste friendly?
- Does it end with next actions?
- Does it avoid sensitive data?
- Is it modular rather than bloated?
