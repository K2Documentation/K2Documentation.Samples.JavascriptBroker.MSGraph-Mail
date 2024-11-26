Migration from Ava to Vitest
===

If you used a previous version of this template, you would probably be using `ava` as the testing framework. whereas now
the template makes use of `vitest`. Below is a simple migration guide on moving from `ava` to `vitest`.

Note: This template has also migrated to `vite` from `babel` and `rollup` directly, so some parts of this will assume
the use of `vite`.

<!-- TOC -->
* [Migration from Ava to Vitest](#migration-from-ava-to-vitest)
    * [1. Installation](#1-installation)
    * [2. Configuring Vitest](#2-configuring-vitest)
    * [3. Syntax Differences](#3-syntax-differences)
      * [Basic Test Structure](#basic-test-structure)
        * [Ava](#ava)
        * [Vitest](#vitest)
      * [Assertions](#assertions)
      * [Test Hooks](#test-hooks)
      * [Example:](#example)
        * [Ava](#ava-1)
        * [Vitest](#vitest-1)
    * [4. Async Tests](#4-async-tests)
        * [Ava](#ava-2)
        * [Vitest](#vitest-2)
    * [5. Snapshot Testing](#5-snapshot-testing)
        * [Vitest](#vitest-3)
    * [6. Coverage](#6-coverage)
    * [7. Running Tests](#7-running-tests)
    * [8. Notes](#8-notes)
<!-- TOC -->

### 1. Installation

First, remove Ava if you haven’t already:

~~~bash
npm uninstall ava
~~~

Then, install Vitest:

~~~bash
npm install -D vitest
~~~

If you’re using Vite, like this template, you can also integrate Vitest directly with it:

~~~bash
npm install -D vitest vite
~~~

### 2. Configuring Vitest

Vitest configuration is typically handled in `vite.config.js` or `vite.config.ts`. However, if you prefer a separate
configuration file, you can create a `vitest.config.ts` file.

Here's an example of `vite.config.ts` with Vitest setup:

```typescript
// vite.config.ts
import {defineConfig} from 'vite';
import {configDefaults} from 'vitest/config';

export default defineConfig({
    test: {
        globals: true, // Ava's globals are auto-enabled, but for Vitest, specify it here
        environment: 'node', // Similar to Ava’s node environment
        exclude: [...configDefaults.exclude, 'dist'], // Customize excludes as needed
        setupFiles: './vitest.setup.ts', // If you have global setup similar to Ava’s hooks
        coverage: {
            provider: 'istanbul' // This template uses istanbul for coverage
        },
    },
});
```

### 3. Syntax Differences

#### Basic Test Structure

Ava uses `test` as the default import, while Vitest lets you import directly from `vitest` with a similar API. Here’s a
comparison:

##### Ava

```javascript
import test from 'ava';

test('basic test', t => {
    t.is(1 + 1, 2);
});
```

##### Vitest

```javascript
import {test, expect} from 'vitest';

test('basic test', () => {
    expect(1 + 1).toBe(2);
});
```

#### Assertions

Ava uses its own assertion library, while Vitest uses Jest-like assertions by default.

| Ava         | Vitest                 |
|-------------|------------------------|
| t.is(a, b)  | expect(a).toBe(b)      |
| t.deepEqual | expect(a).toEqual(b)   |
| t.truthy    | expect(a).toBeTruthy() |
| t.false     | expect(a).toBeFalsy()  |

You’ll need to replace Ava assertions with Vitest’s expect.

#### Test Hooks

Both Ava and Vitest have setup and teardown hooks, but with slightly different names.

| Ava             | Vitest     |
|-----------------|------------|
| test.before     | beforeAll  |
| test.beforeEach | beforeEach |
| test.after      | afterAll   |
| test.afterEach  | afterEach  |

#### Example:

##### Ava

```javascript
test.before(() => {
// runs before all tests
});

test.afterEach(() => {
// runs after each test
});
```

##### Vitest

```javascript
import {beforeAll, afterEach} from 'vitest';

beforeAll(() => {
// runs before all tests
});
afterEach(() => {
// runs after each test
});
```

### 4. Async Tests

Ava allows async tests with `async` functions, and Vitest works the same way but uses a simpler `expect` pattern.

##### Ava

```javascript
test('async test', async t => {
    const data = await fetchData();
    t.is(data, 'expected value');
});
```

##### Vitest

```javascript
test('async test', async () => {
    const data = await fetchData();
    expect(data).toBe('expected value');
});
```

### 5. Snapshot Testing

Vitest has built-in support for snapshot testing, similar to Jest. Ava doesn’t have built-in support for snapshots,
so if you were using a separate library for it, you can replace it with Vitest’s `expect().toMatchSnapshot()`.

##### Vitest

```javascript
test('snapshot test', () => {
    const obj = {foo: 'bar'};
    expect(obj).toMatchSnapshot();
});
```

### 6. Coverage

Vitest integrates with `c8` for coverage, while Ava requires external tools like `nyc`. Vite can also use
`nyc (istanbul)`
as shown in the example `vitest.config.ts`

```bash
npm install -D @vitest/coverage-istanbul
```

Run tests with coverage:

```bash
vitest --coverage
```

Coverage reports will be generated in the `coverage` directory by default.

### 7. Running Tests

- In **Ava**, you might have used `npx ava` or `"test": "ava"`.
- In **Vitest**, you’ll run `vitest` or add it to your package.json:

```json5
{
  //... other config
  "scripts": {
    "test": "vitest --run",
    // Vitest will run with --watch by default, so --run forces a single run and terminate
    "test:watch": "vitest"
  }
}
```

To integrate with Vite, you can also run:

```bash
vite test
```

### 8. Notes

There are some key differences between `ava` and `vitest`:

- Previously, `ava` would run the tests once they were output
  to the `dist` directory as `.js` files, while `vitest` will run the tests as `.ts` file in the `src` directory.
- `vitest` by default expects that tests are named as `*.test.ts` or `*.spec.ts`. This can be changed, but we prefer to
  keep to the defaults for familiarity.

---