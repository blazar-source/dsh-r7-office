# Development Guide

## Prerequisites
- Node.js >= 20.0.0
- (Optional) R7-Office Desktop installed on the host for native conversion and desktop bridge testing.

## Installation
```bash
git clone https://github.com/your-org/dsh-r7-office.git
cd dsh-r7-office
npm test
```

## Running Tests
```bash
# Run all unit, integration, and e2e tests
npm test

# Run specific test suites
npm run test:unit
npm run test:integration
npm run test:e2e
```

## Running MCP Server directly (stdio mode)
```bash
node src/mcp/cli.js
```

## Installing as DeepSeek Harness Plugin
In your DeepSeek Harness profile or workspace:
```yaml
# In cordis.patch.yml:
- insert:
    - id: r7-office
      name: 'dsh-r7-office'
```
Or install via Harness Plugin Manager:
```bash
dsh plugin install /path/to/dsh-r7-office
```
