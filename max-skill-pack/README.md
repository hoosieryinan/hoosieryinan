# Max Skill Pack

A production-ready personal AI skill system for Max: IT workplace manager, IT governance professional, Microsoft 365 practitioner, project lead, content creator, business builder, learner, and reflective father.

This repository is not a generic profile. It is a modular operating system for future AI agents to help Max produce practical work outputs: executive updates, governance documents, M365 plans, PMO artifacts, stakeholder emails, content scripts, business plans, learning paths, and daily reflections.

## Folder structure

```text
max-skill-pack/
  README.md
  AGENTS.md
  skills/
  memory/
  prompts/
  docs/
```

## How to use

1. Choose the skill that matches the task.
2. Provide rough context, desired output, audience, deadline, and language.
3. Ask the AI to follow the relevant `SKILL.md`.
4. Use `examples.md` for behavior patterns and `templates.md` for copy-paste formats.
5. Keep private data out of prompts unless absolutely required; use placeholders.

## Skill catalog

| Skill | Best for |
|---|---|
| `max-workplace-manager` | Weekly planning, stakeholder handling, meeting prep, management updates |
| `max-it-governance` | Governance frameworks, policies, controls, risks, compliance summaries |
| `max-m365-sharepoint` | Microsoft 365, SharePoint, Teams, Purview, migration, PowerShell, troubleshooting |
| `max-project-pmo` | Scope, timeline, RACI, RAID, decisions, project status, vendor coordination |
| `max-executive-communication` | English/Chinese emails, summaries, escalations, meeting invites, stakeholder replies |
| `max-content-creator` | Xiaohongshu, Douyin, WeChat, YouTube, hooks, scripts, captions |
| `max-business-builder` | 每日爪力, pet products, positioning, exhibition planning, AI customer service |
| `max-learning-coach` | Structured learning for technical, business, English, AI, and management topics |
| `max-daily-reflection` | Daily review, discipline tracking, fatherhood notes, self-improvement writing |

## Recommended prompt pattern

```text
Use [skill-name].
Goal: [What I need]
Audience: [Who will read/use it]
Context: [Background]
Raw notes: [Messy notes]
Output: [Email / plan / checklist / table / script]
Language: [English / Chinese / bilingual]
Constraints: [Deadline, tone, risks, privacy limits]
```

## Privacy rule

Do not store secrets, passwords, API keys, tenant identifiers, certificate details, personal IDs, internal URLs, or confidential company data. Use placeholders and keep the pack reusable.

## Maintenance

Update `memory/` when Max's role, projects, preferences, or content direction changes. Update `skills/` when a recurring workflow becomes stable enough to formalize.
