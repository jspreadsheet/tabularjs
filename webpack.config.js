import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export default (env, argv) => {
    const isProduction = argv.mode === 'production';

    return {
        target: 'web',
        entry: './src/index.js',
        output: {
            path: path.resolve(__dirname, isProduction ? 'dist' : 'public'),
            filename: 'index.js',
            library: {
                name: 'tabularjs',
                type: 'umd',
                export: 'default',
            },
            globalObject: 'this',
            clean: isProduction,
        },
        resolve: {
            extensions: ['.js'],
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
};
