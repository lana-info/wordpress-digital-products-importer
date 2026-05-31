# Project Agent Rules

You are responsible for delivering a working application, not just writing code.

Before saying the task is complete:

1. Inspect the project structure.
2. Identify the correct install, build, lint, test, and run commands.
3. Update or create `README.md` with exact commands.
4. Ask only the minimum needed questions if something essential is missing.
5. Run dependency installation if needed.
6. Implement the requested change.
7. Run a smoke check for the main flow.
8. If the project is complex, enable the advanced layer and run the scenario test.
9. Run lint, typecheck, or build if available.
10. Check browser console and server logs if the project has a UI.
11. Fix all blocking errors.
12. Repeat test/run/fix until the project works.
13. If `TASKS.md` exists, take the next open task from it instead of asking the user to pick one manually.
14. Refactor only when it improves readability, removes duplication, or makes the next task safer, and verify again after the refactor.
15. Treat `README.md` as the source of truth for first-run bootstrap and intake, and keep durable decisions in `README.md` and wiki instead of this agent file.

Definition of done:

- The app or workflow starts successfully.
- The main user scenario works.
- No blocking runtime errors remain.
- Build passes if a build exists.
- `TASKS.md` is updated if it exists in the project.
- `CHANGELOG.md` is updated if it exists in the project.
- Small refactors are allowed when they improve the code without changing behavior.
- The final report lists the exact commands used and the result.

Do not stop after writing code.
Do not ask the user to test manually until you have already run your own verification.
If something cannot be tested locally, explain exactly why and what evidence you checked instead.
