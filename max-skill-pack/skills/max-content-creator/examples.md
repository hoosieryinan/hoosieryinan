# Examples for Max Content Creator

## Example 1: Rough notes to structured output

**Input from Max**
```text
Use max-content-creator. I need to update management. Migration is mostly okay but vendor delayed permission report. Need avoid blaming them too directly. Ask for decision on whether we continue pilot next week.
```

**Expected behavior**
- Produce an executive summary first.
- Use diplomatic wording.
- Surface the decision needed.
- Include risks and next actions.

**Output pattern**
```markdown
## Executive summary
- The migration preparation remains on track overall, with one dependency pending: the vendor permission report.
- The main decision is whether to proceed with the pilot next week using current assumptions or wait for the completed report.

## Decision needed
Approve one of the following options by [Date]:
| Option | Benefit | Risk | Recommendation |
|---|---|---|---|
| Proceed with pilot | Maintains timeline | Some permission exceptions may be discovered later | Recommended if pilot scope is low-risk |
| Wait for report | Reduces uncertainty | Delays pilot | Use only if pilot users are business-critical |
```

## Example 2: Email drafting

**Input from Max**
```text
Write an English email to vendor asking them to confirm ETA and share blockers. Tone polite but firm.
```

**Output pattern**
```markdown
Subject: Request for ETA Confirmation and Open Blockers

Hi [VendorName],

Could you please confirm the expected delivery date for [Deliverable] and let us know if there are any blockers requiring our support?

This information is needed to finalize the next project checkpoint and confirm whether the planned timeline remains feasible.

Best regards,
Max
```

## Example 3: Decision support

**Input from Max**
```text
Compare options and recommend next step. We have limited time, unclear ownership, and management wants a simple answer.
```

**Expected output sections**
1. Situation summary
2. Decision options table
3. Recommended option
4. Risk if not decided
5. Next action message Max can send
