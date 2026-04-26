# Contributing to Office365 MCP Server

## ⚠️ Limited Contribution Policy

**This is a personal project released for public use, not a community-driven project.**

This repository is **not actively seeking contributions**. Issues and pull requests are not monitored regularly.

### What This Means

**For Users:**
- ✅ Use the software freely (MIT License)
- ✅ Report security issues privately (see below)
- ✅ Read comprehensive documentation
- ⚠️ Limited support available

**For Developers:**
- ✅ Fork and modify for your needs
- ✅ Create your own maintained version
- ✅ Share your fork with others
- ❌ Pull requests not actively reviewed
- ❌ Feature requests not accepted
- ❌ Issues not actively monitored

---

## 🍴 Recommended Approach: Fork This Repository

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
