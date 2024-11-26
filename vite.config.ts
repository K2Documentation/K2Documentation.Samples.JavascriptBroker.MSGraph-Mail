import {defineConfig} from 'vite';
import path from 'path';
import {glob} from 'glob';
import terser from '@rollup/plugin-terser'

/** @type {import('vite').UserConfig} */
export default defineConfig(({command, mode}) => ({
    plugins: [],
    build: {
        sourcemap: true,
        emptyOutDir: true,
        outDir: "dist",
        lib: {
            // Define multiple entry points using glob, finds all .ts files recursively, then uses the name to create entrypoints
            entry: glob
                .sync("./src/**/*.ts")
                .filter(file => !file.includes('.test')) // Exclude tests
                .reduce((entries, file) => {
                        const entryName = path
                            .relative('./src', file)
                            .replace(/\.ts$/, '');
                        entries[entryName] = path.resolve(__dirname, file);
                        return entries;
                    }, {}
                ),
        },
        rollupOptions: {
            output: [
                // Regular files
                {
                    entryFileNames: `[name].js`,
                    chunkFileNames: `[name].js`,
                    assetFileNames: `[name].[ext]`,
                    format: 'es'
                },
                // Minified files
                {
                    entryFileNames: `[name].min.js`,
                    chunkFileNames: `[name].js`,
                    assetFileNames: `[name].[ext]`,
                    sourcemap: false,
                    format: 'es',
                    plugins: [terser()]
                },
            ],
        },
    },
    minify: 'esbuild',
    esbuild: {
        minifyWhitespace: true,
        minifyIdentifiers: true,
        minifySyntax: true,
    }
}));