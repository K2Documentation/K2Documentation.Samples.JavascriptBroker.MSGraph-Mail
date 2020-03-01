# K2 TypeScript Broker Template

This is demonstrates best-practices for developing a K2 Broker using TypeScript.

# Features

  - Full object model intellisense for making development easier
  - Sample broker code that accesses jsonplaceholder.
  - Sample unit tests with mocks and code coverage.
  - RollupJS configuration for TypeScript.

## Getting Started

This template requires [Node.js](https://nodejs.org/) v12.14.1+ to run.

Install the dependencies and devDependencies:

```bash
npm install
```

See the documentation for [@k2oss/k2-broker-core](https://www.npmjs.com/package/@k2oss/k2-broker-core)
for more information about how to use the broker SDK package.

## Running Unit Tests
To run the unit tests, run:

```bash
npm test
```

You can also use a development build, for debugging and coverage gutters:

```bash
npm run test:dev
```

You will find the code coverage results in [coverage/index.html](./coverage/index.html).

## Building your bundled JS
When you're ready to build your broker, run the following command

```bash
npm run build
```

You will find the results in the [dist/index.js](./dist/index.js).

## Creating a service type
Once you have a bundled .js file, upload it to your repository (anonymously
accessible) and register the service type using the system SmartObject located
at System > Management > SmartObjects > SmartObjects > JavaScript Service
Provider and run the Create From URL method.

### License

MIT, found in the [LICENSE](./LICENSE) file.

[www.k2.com](https://www.k2.com)
