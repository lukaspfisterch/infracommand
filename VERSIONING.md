# Versioning Guidelines

This project follows [Semantic Versioning](https://semver.org/) (SemVer) for version management.

## Version Format

```
MAJOR.MINOR.PATCH[-PRERELEASE][+BUILD]
```

### Version Components

- **MAJOR**: Incompatible API changes
- **MINOR**: New functionality in a backwards compatible manner
- **PATCH**: Backwards compatible bug fixes
- **PRERELEASE**: Optional pre-release identifier (e.g., `-rc1`, `-beta.1`, `-alpha.1`)
- **BUILD**: Optional build metadata (e.g., `+20250127.1`)

## Examples

- `1.0.0` - Initial stable release
- `1.0.1` - Bug fix release
- `1.1.0` - New feature release
- `2.0.0` - Breaking changes release
- `1.2.0-rc1` - Release candidate
- `1.2.0-beta.1` - Beta release
- `1.2.0-alpha.1` - Alpha release

## Release Process

### 1. Pre-release (Optional)
```bash
# Create and push pre-release tag
git tag v1.2.0-rc1
git push origin v1.2.0-rc1
```

### 2. Stable Release
```bash
# Create and push stable release tag
git tag v1.2.0
git push origin v1.2.0
```

### 3. Hotfix Release
```bash
# Create hotfix from main branch
git checkout main
git pull origin main
# Make hotfix changes
git commit -m "fix: resolve critical issue"
git tag v1.2.1
git push origin v1.2.1
```

## Branch Strategy

- **main**: Production-ready code
- **develop**: Integration branch for features
- **feature/***: Feature development branches
- **hotfix/***: Critical bug fixes

## Commit Message Format

We follow [Conventional Commits](https://www.conventionalcommits.org/):

```
<type>[optional scope]: <description>

[optional body]

[optional footer(s)]
```

### Types
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `style`: Code style changes (formatting, etc.)
- `refactor`: Code refactoring
- `test`: Adding or updating tests
- `chore`: Maintenance tasks

### Examples
```
feat: add window management capabilities
fix: resolve PowerShell execution policy issue
docs: update installation instructions
chore: update dependencies
```

## Automated Release

GitHub Actions automatically creates releases when you push tags:

1. **Pre-release**: Tag with `-rc`, `-beta`, or `-alpha` suffix
2. **Stable release**: Tag without suffix
3. **Hotfix**: Patch version increment

## Version Bumping

### Manual Bumping
```bash
# Update version in pyproject.toml
# Create and push tag
git tag v1.2.0
git push origin v1.2.0
```

### Automatic Bumping
Use tools like `bump2version` or `semantic-release` for automatic version management.

## Changelog

- Update `CHANGELOG.md` for each release
- Follow [Keep a Changelog](https://keepachangelog.com/) format
- Include all changes since last release
