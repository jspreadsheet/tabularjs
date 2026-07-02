import { expect } from 'chai';
import fs from 'fs';

describe('Packaging', () => {
    const pkg = JSON.parse(fs.readFileSync('package.json', 'utf8'));

    it('should declare an exports map', () => {
        expect(pkg.exports).to.be.an('object');
        expect(pkg.exports['.']).to.be.an('object');
    });

    it('should point every exports target at an existing file', () => {
        const targets = [];
        const collect = (value) => {
            if (typeof value === 'string') targets.push(value);
            else if (value && typeof value === 'object') Object.values(value).forEach(collect);
        };
        collect(pkg.exports);
        expect(targets.length).to.be.greaterThan(0);
        for (const target of targets) {
            expect(fs.existsSync(target), `${target} should exist`).to.equal(true);
        }
    });

    it('should resolve the package by name in Node (self-reference, regression)', async () => {
        // The published package could not be imported at all: main pointed at
        // a UMD bundle inside a "type": "module" package, so the UMD wrapper
        // ran as ESM where top-level `this` is undefined and threw.
        const m = await import('tabularjs');
        expect(m.default).to.be.a('function');
    });

    it('should expose a working parser through the Node entry point', async function () {
        if (!fs.existsSync('samples/test1.csv')) {
            this.skip();
        }
        const { default: tabularjs } = await import('tabularjs');
        const result = await tabularjs('samples/test1.csv');
        expect(result.worksheets).to.be.an('array');
        expect(result.worksheets[0].data.length).to.be.greaterThan(0);
    });

    it('should ship an ESM bundle for bundlers/browsers', async () => {
        expect(fs.existsSync('dist/index.mjs'), 'dist/index.mjs should exist (run npm run build)').to.equal(true);
        const m = await import('../dist/index.mjs');
        expect(m.default).to.be.a('function');
    });
});
