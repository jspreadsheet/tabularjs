# Contributing to TabularJS

Thank you for your interest in contributing to TabularJS!

## Development Setup

```bash
# Clone the repository
git clone https://github.com/jspreadsheet/tabularjs.git
cd tabularjs

# Install dependencies
npm install

# Run tests
npm test

# Run tests with coverage
npm run test:coverage

# Run tests in watch mode
npm run test:watch

# Build the project
npm run build

# Start development server
npm start
```

## Running Tests

```bash
# Run all tests
npm test

# Run specific test file
npx mocha test/helpers.test.js

# Run tests with coverage
npm run test:coverage
```

## Code Style

- Use ES6+ features
- Write clear, self-documenting code
- Add JSDoc comments for public APIs
- Keep functions small and focused
- Follow existing code patterns

## Submitting Changes

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass (`npm test`)
6. Commit your changes (`git commit -m 'Add amazing feature'`)
7. Push to the branch (`git push origin feature/amazing-feature`)
8. Open a Pull Request

## Pull Request Guidelines

- Update documentation if needed
- Add tests for new features
- Ensure all tests pass
- Keep PRs focused on a single feature/fix
- Write clear commit messages

## Reporting Issues

- Use the GitHub issue tracker
- Include code examples when possible
- Specify the file format and browser/Node.js version
- Provide sample files if applicable

## License

By contributing, you agree that your contributions will be licensed under the MIT License.
