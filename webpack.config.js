import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export default (env, argv) => {
    const isProduction = argv.mode === 'production';

    const base = {
        target: 'web',
        entry: './src/index.js',
        resolve: {
            extensions: ['.js'],
            // Node-only optional dependencies: browsers rely on TextDecoder,
            // so keep these out of the browser bundles (~400 KiB of tables)
            alias: {
                'iconv-lite': false,
                'chardet': false,
            },
            fallback: {
                "fs": false,
                "util": false,
                "path": false,
                "stream": false,
                "buffer": false,
            },
        },
        optimization: {
            minimize: isProduction,
            splitChunks: false,
            runtimeChunk: false,
        },
        module: {
            parser: {
                javascript: {
                    dynamicImportMode: 'eager',
                },
            },
        },
        performance: {
            hints: false,
        },
        mode: argv.mode || 'development',
        devtool: false,
        stats: {
            colors: true,
            modules: false,
            children: false,
            chunks: false,
            chunkModules: false
        }
    };

    // UMD bundle for <script> tags and legacy consumers
    const umd = {
        ...base,
        output: {
            path: path.resolve(__dirname, isProduction ? 'dist' : 'public'),
            filename: 'index.js',
            library: {
                name: 'tabularjs',
                type: 'umd',
                export: 'default',
            },
            // 'this' is undefined at the top level of ES modules, which made
            // the UMD wrapper throw when loaded via import
            globalObject: 'globalThis',
            // Cleaning is done by the build script: with clean here, the two
            // parallel production configs delete each other's output
            clean: false,
        },
    };

    if (!isProduction) {
        return umd;
    }

    // Real ES module bundle for bundlers and browser import
    const esm = {
        ...base,
        experiments: {
            outputModule: true,
        },
        output: {
            path: path.resolve(__dirname, 'dist'),
            filename: 'index.mjs',
            library: {
                type: 'module',
            },
            clean: false,
        },
    };

    return [umd, esm];
};
