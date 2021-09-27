<!-- omit in toc -->
# <a href="https://github.com/GruberMarkus/Set-OutlookSignatures" target="_blank"><img src="../src/logo/Set-OutlookSignatures%20Logo.png" width="400" title="Set-OutlookSignatures" alt="Set-OutlookSignatures"></a><br>Centrally manage and deploy Outlook text signatures and Out of Office auto reply messages.<br><a href="https://github.com/GruberMarkus/Set-OutlookSignatures/blob/main/license.txt" target="_blank"><img src="https://img.shields.io/github/license/GruberMarkus/Set-OutlookSignatures" alt=""></a> <a href="https://www.paypal.com/donate?business=JBM584K3L5PX4&no_recurring=0&currency_code=EUR" target="_blank"><img src="https://img.shields.io/badge/sponsor-grey?logo=paypal" alt=""></a> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/views.svg" alt="" data-external="1"> <img src="https://raw.githubusercontent.com/GruberMarkus/my-traffic2badge/traffic/traffic-Set-OutlookSignatures/clones.svg" alt="" data-external="1"> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/network" target="_blank"><img src="https://img.shields.io/github/forks/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/releases" target="_blank"><img src="https://img.shields.io/github/downloads/GruberMarkus/Set-OutlookSignatures/total" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/stargazers" target="_blank"><img src="https://img.shields.io/github/stars/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a> <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank"><img src="https://img.shields.io/github/issues/GruberMarkus/Set-OutlookSignatures" alt="" data-external="1"></a>  

# Welcome!
Thank you very much for your interest in Set-OutlookSignatures.

If you have an idea for a new feature or have found a problem, please <a href="https://github.com/GruberMarkus/Set-OutlookSignatures/issues" target="_blank">create an issue on GitHub</a>.

If you want to contribute code, this document gives you a rough overview of the proposed process.  
I'm not a professional developer - if you are one and you notice something negative in the code or process, please let me know. 
# Branches
The default branch is named '`main`'. It contains the source of the latest stable release.

Tags in the '`main`' branch mark releases - we release on tags (and therefore commits) and not on branches.
# Development process
1. Create a new branch based on '`main`'
   - Hotfix: Give the branch a name starting with '`hotfix-`', e. g. '`hotfix-1.2.3`' or '`hotfix-issue13`'
   - New feature or vNext: Give the branch a name starting with '`develop-'`, e. g. '`develop-1.3.0`' or '`develop-vNext`'
2. Work on the code in the new branch. Every commit to the GitHub repository will trigger the build process and create a draft release.  
You can commit to the new branch as often as you like, and you don't have to care about commit messages during development and testing - the commit messages from the dev branch will not appear in the '`main`' branch, as we will do a "squash and merge" later.
3. When development is done, update the changelog.<br>Don't forget the link to the new release at the end of the file.
4. Create a pull request to incorporate the '`hotfix-`' or '`develop-`' changes into the '`main`' branch.
5. Discuss and review the pull request.
6. When applying the pull request, use `"squash and merge"` as it helps keep the commit history in the '`main`' branch clean and allows the developer to focus on, well, development.
7. After the pull request has been committed to the '`main`' branch, delete the now obsolete '`hotfix-`' or '`develop-`' branch.
8. If there are other '`hotfix-`' or '`develop-`' branches, they have to be rebased to the '`main`' branch which is now at least one commit ahead.
# Commit messages
Commit messages should follow the <a href="https://www.conventionalcommits.org" target="_blank">Conventional Commits</a> standard.
## Commit message format
```
<type>[optional scope]: <short description>
<blank line>
[optional body]
<blank line>
[optional footer]
```
## Type
- Type is mandatory. 
- '`fix[optional scope]:`' A fix. Bumps SemVer patch version.
- '`feat[optional scope]:`' Introduces a new feature to the codebase. Bumps SemVer minor version.
- Other commit types other than '`fix:`' and '`feat:`' are allowed, e. g. '`build:`', '`chore:`', '`ci:`', '`docs:`', '`perf:`', '`refactor:`', '`revert:`', '`style:`' and '`test:`'.
- A scope may be provided to a commit's type, to provide additional contextual information and is contained within parenthesis, e.g. '`feat(parser): add ability to parse arrays`'.
## Body
- Body is optional.
- Provide additional contextual information about the code changes. The body must begin one blank line after the description.
- '`BREAKING CHANGE:<blank>`' at the beginning of the optional body or footer section introduces a breaking API change (bumps SemVer major version). A breaking change can be part of commits of any type.
## Footer
- Footer is optional.
- The footer should contain additional issue references about the code changes (such as the issues it fixes, e.g. '`Fixes [#13](https://github.com/GruberMarkus/Set-OutlookSignatures/issues/13) ([@GruberMarkus](https://github.com/GruberMarkus))`'.
- Text describing further details.
- '`BREAKING CHANGE:<blank>`' at the beginning of the optional body or footer section introduces a breaking API change (bumpgs SemVer major version). A breaking change can be part of commits of any type.
# Build process
Every single commit in any branch or setting a tag starting with '`v`' triggers a build and the creation of a draft release.

The draft release includes the build artifact(s), the corresponding changelog entry and file hash information.

The build artifacts can be downloaded and go through the final test process.
- If these final tests are passed and the information in the draft build is correct, the build can directly be released.  
- If these final tests fail or the information in the draft release is wrong, delete the draft release and go on with with development process.

The build process is built on GitHub Actions workflows and currently only works in this environment.
