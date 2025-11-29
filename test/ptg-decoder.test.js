import { expect } from 'chai';
import { decodePTG } from '../src/utils/ptg-decoder.js';

describe('PTG Decoder', () => {
    describe('Basic Operands', () => {
        it('should decode integer (tInt)', () => {
            // tInt token: 0x1E, followed by 2-byte integer (little-endian)
            // Example: integer value 42 = 0x2A00
            const tokens = new Uint8Array([0x1E, 0x2A, 0x00]);
            const result = decodePTG(tokens);
            expect(result).to.equal('42');
        });

        it('should decode floating point number (tNum)', () => {
            // tNum token: 0x1F, followed by 8-byte IEEE 754 double (little-endian)
            // Example: 3.14
            const buffer = new ArrayBuffer(9);
            const view = new DataView(buffer);
            view.setUint8(0, 0x1F); // tNum token
            view.setFloat64(1, 3.14, true); // little-endian
            const tokens = new Uint8Array(buffer);
            const result = decodePTG(tokens);
            expect(result).to.equal('3.14');
        });

        it('should decode ASCII string (tStr)', () => {
            // tStr token: 0x17, length byte, unicode flag byte, then string bytes
            // Example: "Test" (ASCII)
            const tokens = new Uint8Array([
                0x17,           // tStr
                0x04,           // length = 4
                0x00,           // not unicode
                0x54, 0x65, 0x73, 0x74  // "Test"
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('"Test"');
        });

        it('should decode unicode string (tStr)', () => {
            // tStr with unicode flag
            // Example: "Hi" in unicode
            const tokens = new Uint8Array([
                0x17,           // tStr
                0x02,           // length = 2
                0x01,           // unicode flag
                0x48, 0x00,     // 'H'
                0x69, 0x00      // 'i'
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('"Hi"');
        });

        it('should decode boolean TRUE (tBool)', () => {
            const tokens = new Uint8Array([0x1D, 0x01]);
            const result = decodePTG(tokens);
            expect(result).to.equal('TRUE');
        });

        it('should decode boolean FALSE (tBool)', () => {
            const tokens = new Uint8Array([0x1D, 0x00]);
            const result = decodePTG(tokens);
            expect(result).to.equal('FALSE');
        });

        it('should decode missing argument (tMissArg)', () => {
            const tokens = new Uint8Array([0x16]);
            const result = decodePTG(tokens);
            expect(result).to.equal('');
        });
    });

    describe('Binary Operators', () => {
        it('should decode addition (tAdd)', () => {
            // Push 2, push 3, add
            const tokens = new Uint8Array([
                0x1E, 0x02, 0x00,  // int 2
                0x1E, 0x03, 0x00,  // int 3
                0x03               // add
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('2+3');
        });

        it('should decode subtraction (tSub)', () => {
            // Push 10, push 5, subtract
            const tokens = new Uint8Array([
                0x1E, 0x0A, 0x00,  // int 10
                0x1E, 0x05, 0x00,  // int 5
                0x04               // subtract
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('10-5');
        });

        it('should decode multiplication (tMul)', () => {
            // Push 4, push 5, multiply
            const tokens = new Uint8Array([
                0x1E, 0x04, 0x00,  // int 4
                0x1E, 0x05, 0x00,  // int 5
                0x05               // multiply
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('4*5');
        });

        it('should decode division (tDiv)', () => {
            // Push 20, push 4, divide
            const tokens = new Uint8Array([
                0x1E, 0x14, 0x00,  // int 20
                0x1E, 0x04, 0x00,  // int 4
                0x06               // divide
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('20/4');
        });

        it('should decode power (tPower)', () => {
            // Push 2, push 3, power
            const tokens = new Uint8Array([
                0x1E, 0x02, 0x00,  // int 2
                0x1E, 0x03, 0x00,  // int 3
                0x07               // power
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('2^3');
        });

        it('should decode concatenation (tConcat)', () => {
            // Push "Hello", push "World", concat
            const tokens = new Uint8Array([
                0x17, 0x05, 0x00, 0x48, 0x65, 0x6C, 0x6C, 0x6F,  // "Hello"
                0x17, 0x05, 0x00, 0x57, 0x6F, 0x72, 0x6C, 0x64,  // "World"
                0x08  // concat
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('"Hello"&"World"');
        });
    });

    describe('Comparison Operators', () => {
        it('should decode less than (tLT)', () => {
            const tokens = new Uint8Array([
                0x1E, 0x05, 0x00,  // int 5
                0x1E, 0x0A, 0x00,  // int 10
                0x09               // less than
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('5<10');
        });

        it('should decode less than or equal (tLE)', () => {
            const tokens = new Uint8Array([
                0x1E, 0x05, 0x00,  // int 5
                0x1E, 0x05, 0x00,  // int 5
                0x0A               // less than or equal
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('5<=5');
        });

        it('should decode equal (tEQ)', () => {
            const tokens = new Uint8Array([
                0x1E, 0x05, 0x00,  // int 5
                0x1E, 0x05, 0x00,  // int 5
                0x0B               // equal
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('5=5');
        });

        it('should decode greater than or equal (tGE)', () => {
            const tokens = new Uint8Array([
                0x1E, 0x0A, 0x00,  // int 10
                0x1E, 0x05, 0x00,  // int 5
                0x0C               // greater than or equal
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('10>=5');
        });

        it('should decode greater than (tGT)', () => {
            const tokens = new Uint8Array([
                0x1E, 0x0A, 0x00,  // int 10
                0x1E, 0x05, 0x00,  // int 5
                0x0D               // greater than
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('10>5');
        });

        it('should decode not equal (tNE)', () => {
            const tokens = new Uint8Array([
                0x1E, 0x05, 0x00,  // int 5
                0x1E, 0x0A, 0x00,  // int 10
                0x0E               // not equal
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('5<>10');
        });
    });

    describe('Unary Operators', () => {
        it('should decode unary plus (tUplus)', () => {
            const tokens = new Uint8Array([
                0x1E, 0x05, 0x00,  // int 5
                0x12               // unary plus
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('+5');
        });

        it('should decode unary minus (tUminus)', () => {
            const tokens = new Uint8Array([
                0x1E, 0x05, 0x00,  // int 5
                0x13               // unary minus
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('-5');
        });

        it('should decode percent (tPercent)', () => {
            const tokens = new Uint8Array([
                0x1E, 0x32, 0x00,  // int 50
                0x14               // percent
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('50%');
        });

        it('should decode parentheses (tParen)', () => {
            const tokens = new Uint8Array([
                0x1E, 0x02, 0x00,  // int 2
                0x1E, 0x03, 0x00,  // int 3
                0x03,              // add
                0x15               // paren
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('(2+3)');
        });
    });

    describe('Cell References', () => {
        it('should decode relative cell reference (tRef)', () => {
            // tRef: 0x24, 2 bytes row, 2 bytes col with flags
            // Cell A1 (row=0, col=0, both relative)
            const tokens = new Uint8Array([
                0x24,           // tRef
                0x00, 0x00,     // row 0
                0x00, 0xC0      // col 0 with relative flags (0xC000 = both relative)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('A1');
        });

        it('should decode absolute cell reference (tRef)', () => {
            // Cell $A$1 (row=0, col=0, both absolute)
            const tokens = new Uint8Array([
                0x24,           // tRef
                0x00, 0x00,     // row 0
                0x00, 0x00      // col 0 with no relative flags (absolute)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('$A$1');
        });

        it('should decode mixed cell reference (tRef)', () => {
            // Cell A$1 (row absolute, col relative)
            const tokens = new Uint8Array([
                0x24,           // tRef
                0x00, 0x00,     // row 0
                0x00, 0x40      // col 0, col relative (0x4000)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('A$1');
        });

        it('should decode cell reference with larger coordinates', () => {
            // Cell B5 (row=4, col=1, both relative)
            const tokens = new Uint8Array([
                0x24,           // tRef
                0x04, 0x00,     // row 4
                0x01, 0xC0      // col 1 with relative flags
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('B5');
        });
    });

    describe('Area References (Ranges)', () => {
        it('should decode relative range reference (tArea)', () => {
            // tArea: 0x25, 2 bytes row1, 2 bytes row2, 2 bytes col1, 2 bytes col2
            // Range A1:B2 (all relative)
            const tokens = new Uint8Array([
                0x25,           // tArea
                0x00, 0x00,     // row1 = 0
                0x01, 0x00,     // row2 = 1
                0x00, 0xC0,     // col1 = 0 (relative)
                0x01, 0xC0      // col2 = 1 (relative)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('A1:B2');
        });

        it('should decode absolute range reference (tArea)', () => {
            // Range $A$1:$C$5 (all absolute)
            const tokens = new Uint8Array([
                0x25,           // tArea
                0x00, 0x00,     // row1 = 0
                0x04, 0x00,     // row2 = 4
                0x00, 0x00,     // col1 = 0 (absolute)
                0x02, 0x00      // col2 = 2 (absolute)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('$A$1:$C$5');
        });

        it('should decode mixed range reference (tArea)', () => {
            // Range A$1:$B2 (mixed)
            const tokens = new Uint8Array([
                0x25,           // tArea
                0x00, 0x00,     // row1 = 0
                0x01, 0x00,     // row2 = 1
                0x00, 0x40,     // col1 = 0 (col relative, row absolute)
                0x01, 0x80      // col2 = 1 (col absolute, row relative)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('A$1:$B2');
        });
    });

    describe('Range Operators', () => {
        it('should decode range operator (tRange)', () => {
            // Push A1, push B5, range
            const tokens = new Uint8Array([
                0x24, 0x00, 0x00, 0x00, 0xC0,  // A1
                0x24, 0x04, 0x00, 0x01, 0xC0,  // B5
                0x11                            // range operator
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('A1:B5');
        });

        it('should decode union operator (tUnion)', () => {
            // Push A1, push B2, union
            const tokens = new Uint8Array([
                0x24, 0x00, 0x00, 0x00, 0xC0,  // A1
                0x24, 0x01, 0x00, 0x01, 0xC0,  // B2
                0x10                            // union operator
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('A1,B2');
        });
    });

    describe('Functions', () => {
        it('should decode function with fixed args (tFunc)', () => {
            // PI function (index 19, no args in reality but decoder assumes 1)
            const tokens = new Uint8Array([
                0x1E, 0x05, 0x00,  // int 5 (dummy arg for test)
                0x21,              // tFunc
                0x13, 0x00         // function index 19 (PI)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('PI(5)');
        });

        it('should decode function with variable args (tFuncVar) - SUM', () => {
            // SUM function (index 4) with 2 arguments
            const tokens = new Uint8Array([
                0x1E, 0x0A, 0x00,  // int 10
                0x1E, 0x14, 0x00,  // int 20
                0x22,              // tFuncVar
                0x02,              // 2 arguments
                0x04, 0x00         // function index 4 (SUM)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('SUM(10,20)');
        });

        it('should decode function with variable args (tFuncVar) - IF', () => {
            // IF function (index 1) with 3 arguments
            const tokens = new Uint8Array([
                0x1D, 0x01,        // TRUE
                0x1E, 0x01, 0x00,  // int 1
                0x1E, 0x00, 0x00,  // int 0
                0x22,              // tFuncVar
                0x03,              // 3 arguments
                0x01, 0x00         // function index 1 (IF)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('IF(TRUE,1,0)');
        });

        it('should decode function with zero args (tFuncVar)', () => {
            // Function with 0 args
            const tokens = new Uint8Array([
                0x22,              // tFuncVar
                0x00,              // 0 arguments
                0x13, 0x00         // function index 19 (PI)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('PI()');
        });
    });

    describe('Complex Formulas', () => {
        it('should decode formula with multiple operations', () => {
            // (2 + 3) * 4
            const tokens = new Uint8Array([
                0x1E, 0x02, 0x00,  // int 2
                0x1E, 0x03, 0x00,  // int 3
                0x03,              // add
                0x15,              // paren
                0x1E, 0x04, 0x00,  // int 4
                0x05               // multiply
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('(2+3)*4');
        });

        it('should decode formula with cell references and operations', () => {
            // A1 + B1
            const tokens = new Uint8Array([
                0x24, 0x00, 0x00, 0x00, 0xC0,  // A1
                0x24, 0x00, 0x00, 0x01, 0xC0,  // B1
                0x03                            // add
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('A1+B1');
        });

        it('should decode formula with function and range', () => {
            // SUM(A1:B2)
            const tokens = new Uint8Array([
                0x25, 0x00, 0x00, 0x01, 0x00, 0x00, 0xC0, 0x01, 0xC0,  // A1:B2
                0x22,              // tFuncVar
                0x01,              // 1 argument
                0x04, 0x00         // function index 4 (SUM)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('SUM(A1:B2)');
        });

        it('should decode nested function formula', () => {
            // IF(TRUE, SUM(1,2), 0)
            const tokens = new Uint8Array([
                0x1D, 0x01,        // TRUE
                0x1E, 0x01, 0x00,  // int 1
                0x1E, 0x02, 0x00,  // int 2
                0x22, 0x02, 0x04, 0x00,  // SUM with 2 args
                0x1E, 0x00, 0x00,  // int 0
                0x22, 0x03, 0x01, 0x00   // IF with 3 args
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('IF(TRUE,SUM(1,2),0)');
        });
    });

    describe('Edge Cases', () => {
        it('should handle empty token array', () => {
            const tokens = new Uint8Array([]);
            const result = decodePTG(tokens);
            expect(result).to.equal('');
        });

        it('should handle null tokens', () => {
            const result = decodePTG(null);
            expect(result).to.equal('');
        });

        it('should handle undefined tokens', () => {
            const result = decodePTG(undefined);
            expect(result).to.equal('');
        });

        it('should handle unknown token types gracefully', () => {
            // Unknown token 0xFF followed by valid token
            const tokens = new Uint8Array([
                0xFF,              // unknown token
                0x1E, 0x05, 0x00   // int 5
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('5');
        });

        it('should handle truncated tokens gracefully', () => {
            // tInt token without data
            const tokens = new Uint8Array([0x1E]);
            const result = decodePTG(tokens);
            expect(result).to.equal('');
        });

        it('should handle stack underflow with default values', () => {
            // Addition operator without enough operands
            const tokens = new Uint8Array([
                0x1E, 0x05, 0x00,  // int 5 (only one operand)
                0x03               // add (needs two operands)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('0+5');
        });
    });

    describe('Token Class Bits', () => {
        it('should handle token with VALUE class (0x20)', () => {
            // tRef with VALUE class
            const tokens = new Uint8Array([
                0x44,           // tRef | VALUE (0x24 | 0x20)
                0x00, 0x00,     // row 0
                0x00, 0xC0      // col 0 (relative)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('A1');
        });

        it('should handle token with REFERENCE class (0x40)', () => {
            // tRef with REFERENCE class
            const tokens = new Uint8Array([
                0x64,           // tRef | REFERENCE (0x24 | 0x40)
                0x00, 0x00,     // row 0
                0x00, 0xC0      // col 0 (relative)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('A1');
        });

        it('should handle token with ARRAY class (0x60)', () => {
            // tRef with ARRAY class
            const tokens = new Uint8Array([
                0x84,           // tRef | ARRAY (0x24 | 0x60)
                0x00, 0x00,     // row 0
                0x00, 0xC0      // col 0 (relative)
            ]);
            const result = decodePTG(tokens);
            expect(result).to.equal('A1');
        });
    });
});
