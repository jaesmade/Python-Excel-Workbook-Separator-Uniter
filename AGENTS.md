# Agent instructions

## Commit meaningful milestones

- Commit your work frequently at meaningful, self-contained milestones. Examples include a completed feature or fix, a useful refactor, or a distinct set of documentation or tests. For a larger task, make separate commits as each important part is completed instead of waiting until the end.
- Before each commit, run checks appropriate to that milestone and review the diff. If a check cannot be run or fails for an unrelated reason, report that limitation rather than claiming it passed.
- Keep commits focused and use the format `feat:<what was achieved>`, with a short, descriptive outcome after the colon (for example, `feat:add worksheet preview` or `feat:preserve source values`).
- Inspect the working tree and index before committing. Preserve pre-existing or unrelated changes, including staged changes; commit only changes attributable to your task. Do not use a blanket staging or commit command that would include someone else's work. If changes cannot be safely separated, leave them untouched and explain the situation.
- Follow the user's explicit instructions and any higher-priority restrictions on Git operations. Do not push, amend, rewrite history, or commit secrets unless explicitly authorized.
- In your final response, report the commits you made and any remaining uncommitted task changes.
