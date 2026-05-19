# Codex Task Prompts for Max

Use these prompts when asking Codex to build, modify, review, or document projects.

## Build a feature
```text
Act as a senior software engineer. Inspect the repository first. Implement [feature] with minimal, maintainable changes. Follow existing patterns. Add or update tests if the repo has a test framework. Do not hard-code secrets. Commit changes and summarize files changed, tests run, and risks.
```

## Refactor safely
```text
Refactor [module/function] to improve readability and maintainability without changing behavior. Identify current behavior first, make small commits if appropriate, and run relevant tests. Explain any assumptions.
```

## Create documentation
```text
Create practical documentation for [system/process]. Include overview, prerequisites, setup, usage, troubleshooting, security notes, and maintenance checklist. Use placeholders for private values.
```

## Debug an issue
```text
Investigate [issue]. Start by locating relevant code and logs. Provide root cause, impact, fix options, recommended fix, validation plan, and rollback plan. Implement the fix only after identifying the likely root cause.
```

## Build an AI workflow
```text
Design and implement an AI-assisted workflow for [use case]. Include input format, prompt structure, output schema, validation, error handling, privacy boundaries, and examples.
```

## Review for production readiness
```text
Review this project for production readiness. Check structure, security, error handling, logging, configuration, tests, documentation, deployment assumptions, and maintainability. Provide prioritized findings and suggested fixes.
```
