# Contributing to m365-mcp-server

Thank you for your interest in contributing! This document provides guidelines for contributing to this project.

## Getting Started

1. Fork the repository
2. Clone your fork: `git clone https://github.com/YOUR_USERNAME/m365-mcp-server.git`
3. Install dependencies: `npm install`
4. Create a branch: `git checkout -b feature/your-feature-name`

## Development Setup

1. Copy `.env.example` to `.env` and fill in your Azure AD credentials
2. Run in development mode: `npm run dev`
3. Run tests: `npm test`
4. Run linting: `npm run lint`

## Pull Request Process

1. Ensure your code passes all tests: `npm test`
2. Ensure your code passes linting: `npm run lint`
3. Ensure TypeScript compiles without errors: `npm run typecheck`
4. Update documentation if needed
5. Create a Pull Request with a clear description of changes

## Code Style

- Use TypeScript for all new code
- Follow the existing code style (enforced by ESLint and Prettier)
- Write meaningful commit messages
- Add tests for new functionality

## Reporting Issues

When reporting issues, please include:
- A clear description of the problem
- Steps to reproduce
- Expected vs actual behavior
- Environment details (Node.js version, OS, etc.)

## Security Issues

Please report security vulnerabilities privately. See [SECURITY.md](SECURITY.md) for details.

## License

By contributing, you agree that your contributions will be licensed under the MIT License.
