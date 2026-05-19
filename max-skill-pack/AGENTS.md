# AGENTS.md — Instructions for AI Agents Working on Max Skill Pack

## Mission
Maintain this repository as a practical, reusable AI Skills system for Max. Prioritize usefulness, structure, privacy, and real-world execution.

## Style rules
- Be clear, structured, professional, practical, and warm.
- Avoid fluffy motivational language unless the target skill explicitly asks for reflective or content writing.
- Prefer templates, checklists, tables, examples, and next actions.
- For English outputs, use business-professional, simple, direct language.
- For Chinese outputs, use natural Chinese that is readable and not overly formal.

## Privacy and security rules
- Never add passwords, API keys, tenant secrets, certificate details, personal IDs, customer confidential data, internal URLs, or exact sensitive infrastructure identifiers.
- Use placeholders like `[TenantName]`, `[ClientName]`, `[VendorName]`, `[ProjectName]`, `[UserGroup]`, `[Date]`, `[TicketID]`.
- When adding examples, keep them fictional or sanitized.

## Skill authoring standards
Each `SKILL.md` must include:
1. Skill name
2. Purpose
3. When to use this skill
4. User context
5. Operating principles
6. Input format expected from Max
7. Output format
8. Decision logic
9. Style and tone rules
10. Things to avoid
11. Example invocations
12. Example outputs

## Change management
- Keep skills modular. Do not overload one skill with unrelated workflows.
- Add examples when adding new behavior.
- Add templates when a workflow is repeated more than twice.
- Update `docs/update-guide.md` if the maintenance process changes.
- Preserve copy-paste friendliness.

## Testing expectations
Before committing changes, run basic checks:
- Verify the expected folder tree exists.
- Verify every skill has `SKILL.md`, `examples.md`, and `templates.md`.
- Verify no obvious secret patterns are present.
