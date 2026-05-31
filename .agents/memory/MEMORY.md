# Memory Index

Shared memory index for all agents. Read this first; then read only the topic files you need.

user/
- [User profile](user/user-profile.md) - solo VBA dev (Bui-Danh), Vietnamese, weighing portfolio/donate vs commercial

project/
- [VBA project direction](project/vba-project-direction.md) - portfolio + free/open-source + donate; Excel-DNA over VSTO if porting; repo cleanup decisions

features/
- [memory-storage](features/memory-storage.md) - memories go in .agents/memory/; overlap-check-and-announce before creating
- [nested-claude-md](features/nested-claude-md.md) - lazy-loaded folder CLAUDE.md; update folder map + root pointer on change

behaviors/
- [config-stays-in-project](behaviors/config-stays-in-project.md) - all config and memory in the project dir, never user home
- [git-commit-policy](behaviors/git-commit-policy.md) - commit subject type-word, LF only, scope split, no --no-verify

rules/
- [vba-conventions](../rules/code/vba-conventions.md) - VBA coding standard (Option Explicit, Left$, Friend, error pattern)
- [doc-style](../rules/doc/doc-style.md) - plain-text agent-config markdown (no bold/em-dash/arrows)
