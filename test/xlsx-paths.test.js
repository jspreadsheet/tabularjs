import { expect } from 'chai';
import { resolveZipPath } from '../src/parser/xlsx.js';

describe('XLSX Relationship Path Resolution', () => {
    it('should resolve a single parent segment', () => {
        expect(resolveZipPath('xl/drawings', '../media/image1.png')).to.equal('xl/media/image1.png');
    });

    it('should resolve multiple parent segments (regression)', () => {
        // The old replace('../', '') only stripped the first occurrence
        expect(resolveZipPath('xl/drawings', '../../customXml/item1.xml')).to.equal('customXml/item1.xml');
    });

    it('should resolve sibling targets relative to the base directory', () => {
        expect(resolveZipPath('xl/drawings', 'charts/chart1.xml')).to.equal('xl/drawings/charts/chart1.xml');
    });

    it('should resolve absolute targets from the package root', () => {
        expect(resolveZipPath('xl/drawings', '/xl/media/image1.png')).to.equal('xl/media/image1.png');
    });

    it('should ignore current-directory segments', () => {
        expect(resolveZipPath('xl/drawings', './media/./image1.png')).to.equal('xl/drawings/media/image1.png');
    });

    it('should not escape the archive root', () => {
        expect(resolveZipPath('xl/drawings', '../../../../etc/passwd')).to.equal('etc/passwd');
    });

    it('should return an empty string for missing targets', () => {
        expect(resolveZipPath('xl/drawings', '')).to.equal('');
        expect(resolveZipPath('xl/drawings', undefined)).to.equal('');
    });
});
