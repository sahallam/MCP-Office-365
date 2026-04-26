# Contributing to Office365 MCP Server

Thank you for your interest in contributing! This project welcomes **selective contributions** with clear guidelines and realistic expectations.

## 📋 Contribution Policy Summary

**Monthly Review Cycle:** Issues and PRs are reviewed on the **1st of each month**

**We Accept:**
- ✅ Bug fixes (with tests and reproduction steps)
- ✅ Security fixes (reviewed with priority)
- ✅ Documentation corrections

**Please Fork Instead:**
- 🍴 New features (additional services, tools)
- 🍴 Enhancements and optimizations
- 🍴 Refactoring without bug fix
- 🍴 Dependency updates (unless security-critical)

**Response Time:**
- Issues: Reviewed monthly, no guaranteed response
- PRs: Reviewed monthly, no guaranteed merge
- Need faster? Fork it!

---

## 🐛 Reporting Bugs

**Before opening an issue:**
1. Check [existing issues](https://github.com/sahallam/MCP-Office-365/issues)
2. Review the [Troubleshooting Guide](README.md#troubleshooting)
3. Try the [latest release](https://github.com/sahallam/MCP-Office-365/releases)

**When opening an issue, include:**
- Clear, descriptive title
- Detailed description of the problem
- Steps to reproduce (be specific!)
- Expected behavior vs actual behavior
- Your environment:
  - Operating system and version
  - Node.js version (`node --version`)
  - Package version
  - Authentication mode (delegated or app-only)
- Error messages or logs (if applicable)
- Screenshots (if relevant)

**Good example:**
```
Title: "Authentication fails with AADSTS700082 after 90 days"

Description:
After using the connector successfully for 90 days, authentication suddenly fails
with error AADSTS700082 (refresh token expired).

Steps to reproduce:
1. Set up delegated auth with .env configuration
2. Authenticate successfully via device code
3. Wait 90 days without using the connector
4. Try to use any tool (e.g., outlook_list_emails)

Expected: Automatic re-authentication prompt
Actual: Error "Authentication failed: AADSTS700082"

Environment:
- OS: macOS 14.2
- Node.js: v20.10.0
- Package: v1.0.0
- Auth mode: delegated

Error log:
[Paste error here]
```

---

## 🔧 Submitting Pull Requests

### Before You Start

1. **Check if it's accepted** - Review the policy above
2. **Open an issue first** - Discuss the bug/fix before coding
3. **Keep it focused** - One bug fix per PR
4. **Read the code standards** - Follow existing patterns

### Accepted PRs: Bug Fixes

**Requirements:**
- [ ] Fixes a specific bug (not enhancement)
- [ ] Includes reproduction steps
- [ ] Adds test case or verification steps
- [ ] Updates documentation if needed
- [ ] Follows existing code style
- [ ] No breaking changes
- [ ] Commits are clear and atomic

**Process:**
1. Fork the repository
2. Create a branch: `bugfix/issue-123-short-description`
3. Write failing test that reproduces the bug
4. Fix the bug
5. Ensure test passes
6. Update documentation if needed
7. Push to your fork
8. Open PR with description linking to issue

### Accepted PRs: Documentation

**Requirements:**
- [ ] Fixes incorrect information
- [ ] Improves clarity
- [ ] Adds missing information
- [ ] Follows existing documentation style

**Minor doc fixes:** Can skip the issue, just submit PR

---

## 💻 Code Standards for Bug Fixes

When submitting bug fix PRs, follow these standards:

### TypeScript
- Use strict type safety (no implicit `any`)
- All functions have explicit return types
- Follow existing code patterns in the repository

### Security
- Validate all user inputs before processing
- Use validation functions in `src/security.ts`
- Follow the input validation pattern:
  ```typescript
  async myFunction(resourceId: string): Promise<Result> {
    // 1. Validate inputs
    validateResourceId(resourceId, 'resource');
    
    // 2. Perform operation
    const result = await this.graphClient.api(`/path/${resourceId}`).get();
    
    // 3. Return result
    return result;
  }
  ```

### Error Handling
- Don't expose sensitive information in error messages
- Use centralized error handling in `index.ts`
- Test error paths

### Documentation
- Update README.md if behavior changes
- Add inline comments only when "why" is non-obvious
- Don't document what code obviously does

### Testing
- Build successfully: `npm run build`
- No TypeScript errors
- Manual testing of the bug fix
- Include test steps in PR description

---

## 🍴 Want to Add Features? Fork This Repository!

If you want to add features, fix bugs, or customize this project:

### 1. Fork the Repository

Click the "Fork" button at the top of this repository to create your own copy.

### 2. Make Your Changes

```bash
# Clone your fork
git clone https://github.com/YOUR-USERNAME/MCP-Office-365.git
cd MCP-Office-365

# Create a branch
git checkout -b feature/your-feature-name

# Make your changes
# ... edit files ...

# Commit and push
git add .
git commit -m "Add your feature"
git push origin feature/your-feature-name
```

### 3. Maintain Your Fork

- Keep your fork public so others can benefit
- Update the README to indicate it's a fork with additional features
- Maintain it independently at your own pace
- Share it with others who might find it useful

### 4. Credit Original Work

The MIT License requires you to include the original copyright notice. Keep the LICENSE file intact.

---

## 🔐 Security Vulnerabilities

**If you discover a security vulnerability, please report it privately.**

- **Do not** open a public issue
- **Do not** submit a pull request
- **Email:** [Contact via GitHub profile or repository maintainer]

Include:
- Description of the vulnerability
- Steps to reproduce
- Potential impact
- Suggested fix (if you have one)

Security issues will be addressed with priority.

---

## 📚 Development Resources (For Your Fork)

If you've forked this repository and want to develop it further, here are some guidelines from the original implementation.

Thank you for your interest in the Office365 MCP Server!

## Getting Started

1. Fork the repository on GitHub
2. Clone your fork locally:
   ```bash
   git clone https://github.com/YOUR-USERNAME/MCP-Office-365.git
   cd MCP-Office-365
   ```
3. Install dependencies:
   ```bash
   npm install
   ```
4. Build the project:
   ```bash
   npm run build
   ```

## Development Workflow

### Setting Up Your Development Environment

1. Create a `.env` file based on `.env.example`
2. Register a Microsoft Entra ID application (see README.md for detailed instructions)
3. Configure the required API permissions

### Making Changes

1. Create a new branch for your feature or bugfix:
   ```bash
   git checkout -b feature/your-feature-name
   ```
2. Make your changes following the code style guidelines below
3. Build and test your changes:
   ```bash
   npm run build
   npm start
   ```
4. Commit your changes with a clear commit message

### Code Style Guidelines

**TypeScript Standards**
- Use TypeScript strict mode (already configured in `tsconfig.json`)
- All functions must have explicit return types
- No implicit `any` types allowed
- Use meaningful variable and function names
- Follow existing code patterns in the repository

**Security Best Practices**
- Always validate user inputs before processing
- Use the validation functions in `src/security.ts`
- Never log sensitive information (tokens, passwords, email content)
- Follow the input validation pattern used throughout the codebase:
  ```typescript
  async myFunction(resourceId: string): Promise<Result> {
    // 1. Validate inputs
    validateResourceId(resourceId, 'resource');

    // 2. Perform operation
    const result = await this.graphClient.api(`/path/${resourceId}`).get();

    // 3. Return result
    return result;
  }
  ```

**File Organization**
- Each Office 365 service has its own module in `src/tools/`
- Authentication logic stays in `src/auth.ts`
- Security utilities stay in `src/security.ts`
- Type definitions go in `src/types.ts`
- Keep modules focused and loosely coupled

### Adding New Features

When adding new Microsoft Graph API functionality:

1. **Choose the appropriate module** in `src/tools/` (e.g., `outlook.ts`, `teams.ts`)
2. **Add the method** to the corresponding class with proper typing
3. **Register the tool** in `src/index.ts`:
   - Add the tool definition with JSON schema
   - Add the handler in the CallTool request handler
4. **Update documentation** in README.md:
   - Add the tool to the "Available Tools" section
   - Add example prompts if applicable
5. **Test thoroughly** with various inputs and edge cases

### Code Quality Checklist

Before submitting your pull request, ensure:

- [ ] Code compiles without errors (`npm run build`)
- [ ] TypeScript strict mode checks pass
- [ ] All user inputs are validated
- [ ] No secrets or credentials are hardcoded
- [ ] Error messages don't expose sensitive information
- [ ] Code follows existing patterns and style
- [ ] Documentation is updated (README.md, code comments)
- [ ] Commit messages are clear and descriptive

## Pull Request Process

1. **Update documentation**: Ensure README.md reflects any changes
2. **Write a clear PR description**:
   - What does this PR do?
   - Why is this change needed?
   - How has it been tested?
3. **Link related issues**: Reference any related GitHub issues
4. **Keep PRs focused**: One feature or bugfix per PR
5. **Respond to feedback**: Be open to suggestions and code review comments

## Reporting Bugs

When reporting bugs, please include:

- Your operating system and version
- Node.js version (`node --version`)
- Steps to reproduce the issue
- Expected behavior
- Actual behavior
- Any error messages or logs
- Authentication mode (delegated or app-only)

## Security Vulnerabilities

**Do not open public issues for security vulnerabilities.**

If you discover a security vulnerability, please email the maintainers directly at the email listed in the repository. We take security seriously and will respond promptly.

## Questions and Support

- For questions about using the server, open a GitHub Discussion
- For bug reports, open a GitHub Issue
- For feature requests, open a GitHub Issue with the "enhancement" label

## Code of Conduct

Be respectful and professional in all interactions. We're all here to build something great together.

## License

By contributing to this project, you agree that your contributions will be licensed under the MIT License.

## Thank You!

Your contributions help make this project better for everyone. We appreciate your time and effort!
