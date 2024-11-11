import {defineConfig} from 'vite';
import path from 'path';
import {glob} from 'glob';

/** @type {import('vite').UserConfig} */
export default defineConfig(({ command, mode }) => ({
    plugins: [],
    build: {
        sourcemap: true,
        emptyOutDir: true,
        outDir: "dist",
        lib: {
            // Define multiple entry points using glob, finds all .ts files recursively, then uses the name to create entrypoints
            entry: glob
                .sync("./src/**/*.ts")
                .reduce((entries, file) => {
                    const entryName = path
                        .relative('./src', file)
                        .replace(/\.ts$/, '');
                    entries[entryName] = path.resolve(__dirname, file);
                    return entries;
                }, {}),
            formats: ['es'],
        },
        rollupOptions: {
            output: {
                entryFileNames: `[name].min.js`,
                chunkFileNames: `[name].js`,
                assetFileNames: `[name].[ext]`,
                dir: 'dist', // Output directory
                format: 'es',
            },
        },
    },
    minify: 'esbuild',
    esbuild: {
        minifyWhitespace: true,
        minifyIdentifiers: true,
        minifySyntax: true,
    }
}));