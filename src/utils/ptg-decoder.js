/**
 * Lightweight BIFF PTG (Parsed Token) Formula Decoder
 * Decodes binary formula tokens to text formulas
 *
 * Based on BIFF8 specification (Excel 97-2003)
 * This is a focused implementation covering ~90% of common formulas
 */

import {
    getColumnName,
    getCellNameFromCoords,
    readUInt16LE,
    readUInt32LE,
    readFloat64LE
} from './helpers.js';

// PTG Token Types (BIFF8)
const PTG = {
    // Basic operand tokens (no class bits, 0x01-0x1F)
    tExp: 0x01,
    tTbl: 0x02,
    tAdd: 0x03,
    tSub: 0x04,
    tMul: 0x05,
    tDiv: 0x06,
    tPower: 0x07,
    tConcat: 0x08,
    tLT: 0x09,
    tLE: 0x0A,
    tEQ: 0x0B,
    tGE: 0x0C,
    tGT: 0x0D,
    tNE: 0x0E,
    tIsect: 0x0F,
    tUnion: 0x10,
    tRange: 0x11,
    tUplus: 0x12,
    tUminus: 0x13,
    tPercent: 0x14,
    tParen: 0x15,
    tMissArg: 0x16,
    tStr: 0x17,
    tAttr: 0x19,       // Attribute token (control flow metadata)
    tErr: 0x1C,
    tBool: 0x1D,
    tInt: 0x1E,
    tNum: 0x1F,

    // Extended tokens with class bits (>= 0x20)
    // These are BASE values (after masking with 0x1F)
    tArray: 0x00,      // Full values: 0x20, 0x40, 0x60
    tFunc: 0x01,       // Full values: 0x21, 0x41, 0x61
    tFuncVar: 0x02,    // Full values: 0x22, 0x42, 0x62
    tName: 0x03,       // Full values: 0x23, 0x43, 0x63
    tRef: 0x04,        // Full values: 0x24, 0x44, 0x64
    tArea: 0x05,       // Full values: 0x25, 0x45, 0x65
    tMemArea: 0x06,    // Full values: 0x26, 0x46, 0x66
    tMemErr: 0x07,     // Full values: 0x27, 0x47, 0x67
    tRefN: 0x0C,       // Full values: 0x2C, 0x4C, 0x6C (relative reference in shared formulas)
    tAreaN: 0x0D,      // Full values: 0x2D, 0x4D, 0x6D (relative area in shared formulas)

    // Class modifiers (bits 5-6)
    VALUE: 0x20,
    REFERENCE: 0x40,
    ARRAY: 0x60
};

// Built-in Excel functions by index
// Complete BIFF8 function table (indices 0x0000 to 0x017B)
// Based on Microsoft [MS-XLS] specification: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/00b5dd7d-51ca-4938-b7b7-483fe0e5933b
// Built-in Excel functions by index
// Complete table from SheetJS xlsx library
// https://github.com/SheetJS/sheetjs
const FUNCTIONS = {
    0: "COUNT",
    1: "IF",
    2: "ISNA",
    3: "ISERROR",
    4: "SUM",
    5: "AVERAGE",
    6: "MIN",
    7: "MAX",
    8: "ROW",
    9: "COLUMN",
    10: "NA",
    11: "NPV",
    12: "STDEV",
    13: "DOLLAR",
    14: "FIXED",
    15: "SIN",
    16: "COS",
    17: "TAN",
    18: "ATAN",
    19: "PI",
    20: "SQRT",
    21: "EXP",
    22: "LN",
    23: "LOG10",
    24: "ABS",
    25: "INT",
    26: "SIGN",
    27: "ROUND",
    28: "LOOKUP",
    29: "INDEX",
    30: "REPT",
    31: "MID",
    32: "LEN",
    33: "VALUE",
    34: "TRUE",
    35: "FALSE",
    36: "AND",
    37: "OR",
    38: "NOT",
    39: "MOD",
    40: "DCOUNT",
    41: "DSUM",
    42: "DAVERAGE",
    43: "DMIN",
    44: "DMAX",
    45: "DSTDEV",
    46: "VAR",
    47: "DVAR",
    48: "TEXT",
    49: "LINEST",
    50: "TREND",
    51: "LOGEST",
    52: "GROWTH",
    53: "GOTO",
    54: "HALT",
    55: "RETURN",
    56: "PV",
    57: "FV",
    58: "NPER",
    59: "PMT",
    60: "RATE",
    61: "MIRR",
    62: "IRR",
    63: "RAND",
    64: "MATCH",
    65: "DATE",
    66: "TIME",
    67: "DAY",
    68: "MONTH",
    69: "YEAR",
    70: "WEEKDAY",
    71: "HOUR",
    72: "MINUTE",
    73: "SECOND",
    74: "NOW",
    75: "AREAS",
    76: "ROWS",
    77: "COLUMNS",
    78: "OFFSET",
    79: "ABSREF",
    80: "RELREF",
    81: "ARGUMENT",
    82: "SEARCH",
    83: "TRANSPOSE",
    84: "ERROR",
    85: "STEP",
    86: "TYPE",
    87: "ECHO",
    88: "SET.NAME",
    89: "CALLER",
    90: "DEREF",
    91: "WINDOWS",
    92: "SERIES",
    93: "DOCUMENTS",
    94: "ACTIVE.CELL",
    95: "SELECTION",
    96: "RESULT",
    97: "ATAN2",
    98: "ASIN",
    99: "ACOS",
    100: "CHOOSE",
    101: "HLOOKUP",
    102: "VLOOKUP",
    103: "LINKS",
    104: "INPUT",
    105: "ISREF",
    106: "GET.FORMULA",
    107: "GET.NAME",
    108: "SET.VALUE",
    109: "LOG",
    110: "EXEC",
    111: "CHAR",
    112: "LOWER",
    113: "UPPER",
    114: "PROPER",
    115: "LEFT",
    116: "RIGHT",
    117: "EXACT",
    118: "TRIM",
    119: "REPLACE",
    120: "SUBSTITUTE",
    121: "CODE",
    122: "NAMES",
    123: "DIRECTORY",
    124: "FIND",
    125: "CELL",
    126: "ISERR",
    127: "ISTEXT",
    128: "ISNUMBER",
    129: "ISBLANK",
    130: "T",
    131: "N",
    132: "FOPEN",
    133: "FCLOSE",
    134: "FSIZE",
    135: "FREADLN",
    136: "FREAD",
    137: "FWRITELN",
    138: "FWRITE",
    139: "FPOS",
    140: "DATEVALUE",
    141: "TIMEVALUE",
    142: "SLN",
    143: "SYD",
    144: "DDB",
    145: "GET.DEF",
    146: "REFTEXT",
    147: "TEXTREF",
    148: "INDIRECT",
    149: "REGISTER",
    150: "CALL",
    151: "ADD.BAR",
    152: "ADD.MENU",
    153: "ADD.COMMAND",
    154: "ENABLE.COMMAND",
    155: "CHECK.COMMAND",
    156: "RENAME.COMMAND",
    157: "SHOW.BAR",
    158: "DELETE.MENU",
    159: "DELETE.COMMAND",
    160: "GET.CHART.ITEM",
    161: "DIALOG.BOX",
    162: "CLEAN",
    163: "MDETERM",
    164: "MINVERSE",
    165: "MMULT",
    166: "FILES",
    167: "IPMT",
    168: "PPMT",
    169: "COUNTA",
    170: "CANCEL.KEY",
    171: "FOR",
    172: "WHILE",
    173: "BREAK",
    174: "NEXT",
    175: "INITIATE",
    176: "REQUEST",
    177: "POKE",
    178: "EXECUTE",
    179: "TERMINATE",
    180: "RESTART",
    181: "HELP",
    182: "GET.BAR",
    183: "PRODUCT",
    184: "FACT",
    185: "GET.CELL",
    186: "GET.WORKSPACE",
    187: "GET.WINDOW",
    188: "GET.DOCUMENT",
    189: "DPRODUCT",
    190: "ISNONTEXT",
    191: "GET.NOTE",
    192: "NOTE",
    193: "STDEVP",
    194: "VARP",
    195: "DSTDEVP",
    196: "DVARP",
    197: "TRUNC",
    198: "ISLOGICAL",
    199: "DCOUNTA",
    200: "DELETE.BAR",
    201: "UNREGISTER",
    204: "USDOLLAR",
    205: "FINDB",
    206: "SEARCHB",
    207: "REPLACEB",
    208: "LEFTB",
    209: "RIGHTB",
    210: "MIDB",
    211: "LENB",
    212: "ROUNDUP",
    213: "ROUNDDOWN",
    214: "ASC",
    215: "DBCS",
    216: "RANK",
    219: "ADDRESS",
    220: "DAYS360",
    221: "TODAY",
    222: "VDB",
    223: "ELSE",
    224: "ELSE.IF",
    225: "END.IF",
    226: "FOR.CELL",
    227: "MEDIAN",
    228: "SUMPRODUCT",
    229: "SINH",
    230: "COSH",
    231: "TANH",
    232: "ASINH",
    233: "ACOSH",
    234: "ATANH",
    235: "DGET",
    236: "CREATE.OBJECT",
    237: "VOLATILE",
    238: "LAST.ERROR",
    239: "CUSTOM.UNDO",
    240: "CUSTOM.REPEAT",
    241: "FORMULA.CONVERT",
    242: "GET.LINK.INFO",
    243: "TEXT.BOX",
    244: "INFO",
    245: "GROUP",
    246: "GET.OBJECT",
    247: "DB",
    248: "PAUSE",
    251: "RESUME",
    252: "FREQUENCY",
    253: "ADD.TOOLBAR",
    254: "DELETE.TOOLBAR",
    255: "User",
    256: "RESET.TOOLBAR",
    257: "EVALUATE",
    258: "GET.TOOLBAR",
    259: "GET.TOOL",
    260: "SPELLING.CHECK",
    261: "ERROR.TYPE",
    262: "APP.TITLE",
    263: "WINDOW.TITLE",
    264: "SAVE.TOOLBAR",
    265: "ENABLE.TOOL",
    266: "PRESS.TOOL",
    267: "REGISTER.ID",
    268: "GET.WORKBOOK",
    269: "AVEDEV",
    270: "BETADIST",
    271: "GAMMALN",
    272: "BETAINV",
    273: "BINOMDIST",
    274: "CHIDIST",
    275: "CHIINV",
    276: "COMBIN",
    277: "CONFIDENCE",
    278: "CRITBINOM",
    279: "EVEN",
    280: "EXPONDIST",
    281: "FDIST",
    282: "FINV",
    283: "FISHER",
    284: "FISHERINV",
    285: "FLOOR",
    286: "GAMMADIST",
    287: "GAMMAINV",
    288: "CEILING",
    289: "HYPGEOMDIST",
    290: "LOGNORMDIST",
    291: "LOGINV",
    292: "NEGBINOMDIST",
    293: "NORMDIST",
    294: "NORMSDIST",
    295: "NORMINV",
    296: "NORMSINV",
    297: "STANDARDIZE",
    298: "ODD",
    299: "PERMUT",
    300: "POISSON",
    301: "TDIST",
    302: "WEIBULL",
    303: "SUMXMY2",
    304: "SUMX2MY2",
    305: "SUMX2PY2",
    306: "CHITEST",
    307: "CORREL",
    308: "COVAR",
    309: "FORECAST",
    310: "FTEST",
    311: "INTERCEPT",
    312: "PEARSON",
    313: "RSQ",
    314: "STEYX",
    315: "SLOPE",
    316: "TTEST",
    317: "PROB",
    318: "DEVSQ",
    319: "GEOMEAN",
    320: "HARMEAN",
    321: "SUMSQ",
    322: "KURT",
    323: "SKEW",
    324: "ZTEST",
    325: "LARGE",
    326: "SMALL",
    327: "QUARTILE",
    328: "PERCENTILE",
    329: "PERCENTRANK",
    330: "MODE",
    331: "TRIMMEAN",
    332: "TINV",
    334: "MOVIE.COMMAND",
    335: "GET.MOVIE",
    336: "CONCATENATE",
    337: "POWER",
    338: "PIVOT.ADD.DATA",
    339: "GET.PIVOT.TABLE",
    340: "GET.PIVOT.FIELD",
    341: "GET.PIVOT.ITEM",
    342: "RADIANS",
    343: "DEGREES",
    344: "SUBTOTAL",
    345: "SUMIF",
    346: "COUNTIF",
    347: "COUNTBLANK",
    348: "SCENARIO.GET",
    349: "OPTIONS.LISTS.GET",
    350: "ISPMT",
    351: "DATEDIF",
    352: "DATESTRING",
    353: "NUMBERSTRING",
    354: "ROMAN",
    355: "OPEN.DIALOG",
    356: "SAVE.DIALOG",
    357: "VIEW.GET",
    358: "GETPIVOTDATA",
    359: "HYPERLINK",
    360: "PHONETIC",
    361: "AVERAGEA",
    362: "MAXA",
    363: "MINA",
    364: "STDEVPA",
    365: "VARPA",
    366: "STDEVA",
    367: "VARA",
    368: "BAHTTEXT",
    369: "THAIDAYOFWEEK",
    370: "THAIDIGIT",
    371: "THAIMONTHOFYEAR",
    372: "THAINUMSOUND",
    373: "THAINUMSTRING",
    374: "THAISTRINGLENGTH",
    375: "ISTHAIDIGIT",
    376: "ROUNDBAHTDOWN",
    377: "ROUNDBAHTUP",
    378: "THAIYEAR",
    379: "RTD",
    380: "CUBEVALUE",
    381: "CUBEMEMBER",
    382: "CUBEMEMBERPROPERTY",
    383: "CUBERANKEDMEMBER",
    384: "HEX2BIN",
    385: "HEX2DEC",
    386: "HEX2OCT",
    387: "DEC2BIN",
    388: "DEC2HEX",
    389: "DEC2OCT",
    390: "OCT2BIN",
    391: "OCT2HEX",
    392: "OCT2DEC",
    393: "BIN2DEC",
    394: "BIN2OCT",
    395: "BIN2HEX",
    396: "IMSUB",
    397: "IMDIV",
    398: "IMPOWER",
    399: "IMABS",
    400: "IMSQRT",
    401: "IMLN",
    402: "IMLOG2",
    403: "IMLOG10",
    404: "IMSIN",
    405: "IMCOS",
    406: "IMEXP",
    407: "IMARGUMENT",
    408: "IMCONJUGATE",
    409: "IMAGINARY",
    410: "IMREAL",
    411: "COMPLEX",
    412: "IMSUM",
    413: "IMPRODUCT",
    414: "SERIESSUM",
    415: "FACTDOUBLE",
    416: "SQRTPI",
    417: "QUOTIENT",
    418: "DELTA",
    419: "GESTEP",
    420: "ISEVEN",
    421: "ISODD",
    422: "MROUND",
    423: "ERF",
    424: "ERFC",
    425: "BESSELJ",
    426: "BESSELK",
    427: "BESSELY",
    428: "BESSELI",
    429: "XIRR",
    430: "XNPV",
    431: "PRICEMAT",
    432: "YIELDMAT",
    433: "INTRATE",
    434: "RECEIVED",
    435: "DISC",
    436: "PRICEDISC",
    437: "YIELDDISC",
    438: "TBILLEQ",
    439: "TBILLPRICE",
    440: "TBILLYIELD",
    441: "PRICE",
    442: "YIELD",
    443: "DOLLARDE",
    444: "DOLLARFR",
    445: "NOMINAL",
    446: "EFFECT",
    447: "CUMPRINC",
    448: "CUMIPMT",
    449: "EDATE",
    450: "EOMONTH",
    451: "YEARFRAC",
    452: "COUPDAYBS",
    453: "COUPDAYS",
    454: "COUPDAYSNC",
    455: "COUPNCD",
    456: "COUPNUM",
    457: "COUPPCD",
    458: "DURATION",
    459: "MDURATION",
    460: "ODDLPRICE",
    461: "ODDLYIELD",
    462: "ODDFPRICE",
    463: "ODDFYIELD",
    464: "RANDBETWEEN",
    465: "WEEKNUM",
    466: "AMORDEGRC",
    467: "AMORLINC",
    468: "CONVERT",
    724: "SHEETJS",
    469: "ACCRINT",
    470: "ACCRINTM",
    471: "WORKDAY",
    472: "NETWORKDAYS",
    473: "GCD",
    474: "MULTINOMIAL",
    475: "LCM",
    476: "FVSCHEDULE",
    477: "CUBEKPIMEMBER",
    478: "CUBESET",
    479: "CUBESETCOUNT",
    480: "IFERROR",
    481: "COUNTIFS",
    482: "SUMIFS",
    483: "AVERAGEIF",
    484: "AVERAGEIFS"
};

// Function argument counts for tFunc (fixed argument functions)
// Complete table from SheetJS xlsx library
// https://github.com/SheetJS/sheetjs
const FUNCTION_ARG_COUNTS = {
    2: 1,
    3: 1,
    10: 0,
    15: 1,
    16: 1,
    17: 1,
    18: 1,
    19: 0,
    20: 1,
    21: 1,
    22: 1,
    23: 1,
    24: 1,
    25: 1,
    26: 1,
    27: 2,
    30: 2,
    31: 3,
    32: 1,
    33: 1,
    34: 0,
    35: 0,
    38: 1,
    39: 2,
    40: 3,
    41: 3,
    42: 3,
    43: 3,
    44: 3,
    45: 3,
    47: 3,
    48: 2,
    53: 1,
    61: 3,
    63: 0,
    65: 3,
    66: 3,
    67: 1,
    68: 1,
    69: 1,
    70: 1,
    71: 1,
    72: 1,
    73: 1,
    74: 0,
    75: 1,
    76: 1,
    77: 1,
    79: 2,
    80: 2,
    83: 1,
    85: 0,
    86: 1,
    89: 0,
    90: 1,
    94: 0,
    95: 0,
    97: 2,
    98: 1,
    99: 1,
    101: 3,
    102: 3,
    105: 1,
    106: 1,
    108: 2,
    111: 1,
    112: 1,
    113: 1,
    114: 1,
    117: 2,
    118: 1,
    119: 4,
    121: 1,
    126: 1,
    127: 1,
    128: 1,
    129: 1,
    130: 1,
    131: 1,
    133: 1,
    134: 1,
    135: 1,
    136: 2,
    137: 2,
    138: 2,
    140: 1,
    141: 1,
    142: 3,
    143: 4,
    144: 4,
    161: 1,
    162: 1,
    163: 1,
    164: 1,
    165: 2,
    172: 1,
    175: 2,
    176: 2,
    177: 3,
    178: 2,
    179: 1,
    184: 1,
    186: 1,
    189: 3,
    190: 1,
    195: 3,
    196: 3,
    197: 1,
    198: 1,
    199: 3,
    201: 1,
    207: 4,
    210: 3,
    211: 1,
    212: 2,
    213: 2,
    214: 1,
    215: 1,
    225: 0,
    229: 1,
    230: 1,
    231: 1,
    232: 1,
    233: 1,
    234: 1,
    235: 3,
    244: 1,
    247: 4,
    252: 2,
    257: 1,
    261: 1,
    271: 1,
    273: 4,
    274: 2,
    275: 2,
    276: 2,
    277: 3,
    278: 3,
    279: 1,
    280: 3,
    281: 3,
    282: 3,
    283: 1,
    284: 1,
    285: 2,
    286: 4,
    287: 3,
    288: 2,
    289: 4,
    290: 3,
    291: 3,
    292: 3,
    293: 4,
    294: 1,
    295: 3,
    296: 1,
    297: 3,
    298: 1,
    299: 2,
    300: 3,
    301: 3,
    302: 4,
    303: 2,
    304: 2,
    305: 2,
    306: 2,
    307: 2,
    308: 2,
    309: 3,
    310: 2,
    311: 2,
    312: 2,
    313: 2,
    314: 2,
    315: 2,
    316: 4,
    325: 2,
    326: 2,
    327: 2,
    328: 2,
    331: 2,
    332: 2,
    337: 2,
    342: 1,
    343: 1,
    346: 2,
    347: 1,
    350: 4,
    351: 3,
    352: 1,
    353: 2,
    360: 1,
    368: 1,
    369: 1,
    370: 1,
    371: 1,
    372: 1,
    373: 1,
    374: 1,
    375: 1,
    376: 1,
    377: 1,
    378: 1,
    382: 3,
    385: 1,
    392: 1,
    393: 1,
    396: 2,
    397: 2,
    398: 2,
    399: 1,
    400: 1,
    401: 1,
    402: 1,
    403: 1,
    404: 1,
    405: 1,
    406: 1,
    407: 1,
    408: 1,
    409: 1,
    410: 1,
    414: 4,
    415: 1,
    416: 1,
    417: 2,
    420: 1,
    421: 1,
    422: 2,
    424: 1,
    425: 2,
    426: 2,
    427: 2,
    428: 2,
    430: 3,
    438: 3,
    439: 3,
    440: 3,
    443: 2,
    444: 2,
    445: 2,
    446: 2,
    447: 6,
    448: 6,
    449: 2,
    450: 2,
    464: 2,
    468: 3,
    476: 2,
    479: 1,
    480: 2,
    65535: 0
};

FUNCTION_ARG_COUNTS[221] = 0; // TODAY

const ERRORS = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A' };

// Built-in defined names (Lbl.fBuiltin): the name is a single character code
const BUILTIN_NAMES = ['Consolidate_Area', 'Auto_Open', 'Auto_Close', 'Extract', 'Database', 'Criteria', 'Print_Area',
    'Print_Titles', 'Recorder', 'Data_Form', 'Auto_Activate', 'Auto_Deactivate', 'Sheet_Title', '_FilterDatabase'];

/**
 * Thrown for a token stream the decoder cannot render faithfully (unknown token, truncated operand, data table).
 * A wrong formula is worse than no formula, so the caller must not emit a guess.
 */
export class PTGError extends Error {}

/**
 * Convert column number to letter with absolute reference if needed
 */
function formatColRef(col, isAbsolute) {
    const letter = getColumnName(col);
    return isAbsolute ? '$' + letter : letter;
}

/**
 * Format row number with absolute reference if needed
 */
function formatRowRef(row, isAbsolute) {
    return isAbsolute ? '$' + (row + 1) : (row + 1).toString();
}

/**
 * Decode a BIFF8 RgceLoc operand (tRef/tArea): 16-bit absolute row and
 * 14-bit absolute column. The relative flags (bits 15/14 of the column
 * field) only control $ display - the stored coordinates are absolute.
 */
function formatLoc(rowField, colField) {
    const rowRel = (colField & 0x8000) !== 0;
    const colRel = (colField & 0x4000) !== 0;
    return formatColRef(colField & 0x3FFF, !colRel) + formatRowRef(rowField, !rowRel);
}

/**
 * Decode a BIFF8 RgceLocRel operand (tRefN/tAreaN): when a relative flag is
 * set, the field is a signed offset from the host cell (row: 16-bit two's
 * complement, column: 8-bit two's complement); otherwise it is absolute.
 */
function formatLocRel(rowField, colField, cellContext) {
    const rowRel = (colField & 0x8000) !== 0;
    const colRel = (colField & 0x4000) !== 0;
    let rowNum = rowField;
    let colNum = colField & 0x3FFF;

    if (rowRel) {
        if (rowNum & 0x8000) rowNum -= 0x10000; // Sign-extend 16-bit offset
        if (cellContext) rowNum = (rowNum + cellContext.row + 0x10000) % 0x10000;
    }
    if (colRel) {
        colNum = colField & 0xFF;
        if (colNum & 0x80) colNum -= 0x100; // Sign-extend 8-bit offset
        if (cellContext) colNum = (colNum + cellContext.col + 0x100) % 0x100;
    }

    return formatColRef(colNum, !colRel) + formatRowRef(rowNum, !rowRel);
}

/**
 * Quote a sheet name for a formula when it is not a plain identifier ('My Sheet'!A1)
 */
function quoteSheet(name) {
    return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
        ? name : `'${name.replace(/'/g, "''")}'`;
}

/**
 * Sheet prefix of a 3-D reference: ixti indexes EXTERNSHEET, whose entry points to a SUPBOOK
 */
function sheetPrefix(ixti, ctx) {
    const xti = ctx && ctx.externSheets && ctx.externSheets[ixti];
    if (!xti) {
        throw new PTGError(`3-D reference to unknown EXTERNSHEET entry ${ixti}`);
    }
    const book = ctx.supbooks && ctx.supbooks[xti.supbook];
    if (xti.first === 0xFFFF || xti.first === 0xFFFE) {
        return null; // deleted sheet / workbook-level
    }
    let names;
    let prefix = '';
    if (!book || book.type === 'self') {
        names = ctx.sheets || [];
    } else {
        names = book.sheetNames || [];
        prefix = `[${book.name}]`;
    }
    const first = names[xti.first];
    const last = names[xti.last];
    if (first === undefined) {
        throw new PTGError(`3-D reference to unknown sheet ${xti.first}`);
    }
    const label = prefix + first + (xti.last !== xti.first && last !== undefined ? ':' + last : '');
    return quoteSheet(label) + '!';
}

/**
 * Read one constant of a tArray (SerAr in RgbExtra)
 */
function readArrayValue(extra, st) {
    const type = extra[st.pos];
    st.pos++;
    switch (type) {
        case 0x00: st.pos += 8; return '';
        case 0x01: { const v = readFloat64LE(extra, st.pos); st.pos += 8; return v.toString(); }
        case 0x02: {
            const cch = readUInt16LE(extra, st.pos);
            const unicode = extra[st.pos + 2] & 0x01;
            st.pos += 3;
            let s = '';
            for (let i = 0; i < cch; i++) {
                s += String.fromCharCode(unicode ? readUInt16LE(extra, st.pos) : extra[st.pos]);
                st.pos += unicode ? 2 : 1;
            }
            return `"${s.replace(/"/g, '""')}"`;
        }
        case 0x04: { const v = extra[st.pos]; st.pos += 8; return v ? 'TRUE' : 'FALSE'; }
        case 0x10: { const v = extra[st.pos]; st.pos += 8; return ERRORS[v] || '#N/A'; }
        default: throw new PTGError(`unknown array constant type ${type}`);
    }
}

/**
 * Decode BIFF PTG tokens to formula string
 * @param {Uint8Array} tokens - Binary formula tokens (rgce)
 * @param {boolean} debug - Enable debug logging
 * @param {Object} cellContext - Host cell, used by the relative references of shared formulas
 * @param {number} cellContext.row - Target cell row (0-based)
 * @param {number} cellContext.col - Target cell column (0-based)
 * @param {Object} [ctx] - Workbook context: sheets (names), externSheets [{supbook, first, last}],
 *     supbooks [{type: 'self'|'addin'|'external', name, sheetNames, externNames}], names [{name}],
 *     extra (Uint8Array, the RgbExtra after rgce, holding tArray constants)
 * @returns {string} Formula string (without leading =)
 * @throws {PTGError} When the token stream cannot be decoded faithfully
 */
export function decodePTG(tokens, debug = false, cellContext = null, ctx = null) {
    if (!tokens || tokens.length === 0) {
        return '';
    }

    const stack = [];
    const extraState = { pos: 0 };
    let pos = 0;

    const need = (n) => {
        if (pos + n > tokens.length) {
            throw new PTGError(`truncated token at ${pos - 1}`);
        }
    };
    const pop = () => {
        if (!stack.length) {
            throw new PTGError('operand stack underflow');
        }
        return stack.pop();
    };
    const binary = (op) => {
        const b = pop();
        const a = pop();
        stack.push(`${a}${op}${b}`);
    };
    const callArgs = (n) => {
        const args = [];
        for (let i = 0; i < n; i++) {
            args.unshift(pop());
        }
        return args;
    };

    while (pos < tokens.length) {
        const rawToken = tokens[pos];

        if (debug) {
            console.log(`Pos ${pos}: Token 0x${rawToken.toString(16).padStart(2, '0')} Stack before:`, [...stack]);
        }

        pos++;

        if (rawToken >= 0x20) {
            // Operand / function tokens: bits 5-6 are the class (reference / value / array), bits 0-4 the type
            const base = rawToken & 0x1F;
            switch (base) {
                case 0x00: { // tArray: 7 unused bytes, the constants are in RgbExtra
                    need(7);
                    pos += 7;
                    const extra = ctx && ctx.extra;
                    if (!extra || extraState.pos + 3 > extra.length) {
                        throw new PTGError('array constant without RgbExtra');
                    }
                    const cols = extra[extraState.pos] + 1;
                    const rows = readUInt16LE(extra, extraState.pos + 1) + 1;
                    extraState.pos += 3;
                    const lines = [];
                    for (let r = 0; r < rows; r++) {
                        const vals = [];
                        for (let c = 0; c < cols; c++) {
                            vals.push(readArrayValue(extra, extraState));
                        }
                        lines.push(vals.join(','));
                    }
                    stack.push(`{${lines.join(';')}}`);
                    break;
                }
                case 0x01: { // tFunc: fixed argument count
                    need(2);
                    const funcIndex = readUInt16LE(tokens, pos);
                    pos += 2;
                    const funcName = FUNCTIONS[funcIndex];
                    const argCount = FUNCTION_ARG_COUNTS[funcIndex];
                    if (!funcName || argCount === undefined) {
                        throw new PTGError(`unknown fixed-argument function ${funcIndex}`);
                    }
                    stack.push(`${funcName}(${callArgs(argCount).join(',')})`);
                    break;
                }
                case 0x02: { // tFuncVar: argument count + function index
                    need(3);
                    const argCount = tokens[pos] & 0x7F;
                    const funcIndex = readUInt16LE(tokens, pos + 1) & 0x7FFF;
                    pos += 3;
                    const args = callArgs(argCount);
                    if (funcIndex === 255) {
                        // User-defined / add-in function: the first argument is its name (tName / tNameX)
                        const name = args.shift();
                        stack.push(`${name}(${args.join(',')})`);
                    } else {
                        const funcName = FUNCTIONS[funcIndex];
                        if (!funcName) {
                            throw new PTGError(`unknown function ${funcIndex}`);
                        }
                        stack.push(`${funcName}(${args.join(',')})`);
                    }
                    break;
                }
                case 0x03: { // tName: defined name index (1-based) + 2 unused bytes
                    need(4);
                    const idx = readUInt16LE(tokens, pos);
                    pos += 4;
                    const name = ctx && ctx.names && ctx.names[idx - 1];
                    if (!name) {
                        throw new PTGError(`unknown defined name ${idx}`);
                    }
                    stack.push(name.name);
                    break;
                }
                case 0x04: // tRef
                    need(4);
                    stack.push(formatLoc(readUInt16LE(tokens, pos), readUInt16LE(tokens, pos + 2)));
                    pos += 4;
                    break;
                case 0x05: // tArea
                    need(8);
                    stack.push(`${formatLoc(readUInt16LE(tokens, pos), readUInt16LE(tokens, pos + 4))}:` +
                        `${formatLoc(readUInt16LE(tokens, pos + 2), readUInt16LE(tokens, pos + 6))}`);
                    pos += 8;
                    break;
                case 0x06: // tMemArea: 4 unused bytes + cce; the sub-expression tokens follow and are decoded normally
                case 0x07: // tMemErr
                case 0x08: // tMemNoMem
                    need(6);
                    pos += 6;
                    break;
                case 0x09: // tMemFunc: cce
                case 0x0E: // tMemAreaN
                case 0x0F: // tMemNoMemN
                    need(2);
                    pos += 2;
                    break;
                case 0x0A: // tRefErr
                    need(4);
                    pos += 4;
                    stack.push('#REF!');
                    break;
                case 0x0B: // tAreaErr
                    need(8);
                    pos += 8;
                    stack.push('#REF!');
                    break;
                case 0x0C: // tRefN: relative to the host cell (shared formulas)
                    need(4);
                    stack.push(formatLocRel(readUInt16LE(tokens, pos), readUInt16LE(tokens, pos + 2), cellContext));
                    pos += 4;
                    break;
                case 0x0D: // tAreaN
                    need(8);
                    stack.push(`${formatLocRel(readUInt16LE(tokens, pos), readUInt16LE(tokens, pos + 4), cellContext)}:` +
                        `${formatLocRel(readUInt16LE(tokens, pos + 2), readUInt16LE(tokens, pos + 6), cellContext)}`);
                    pos += 8;
                    break;
                case 0x19: { // tNameX: ixti + external name index (1-based) + 2 unused bytes
                    need(6);
                    const ixti = readUInt16LE(tokens, pos);
                    const idx = readUInt16LE(tokens, pos + 2);
                    pos += 6;
                    const xti = ctx && ctx.externSheets && ctx.externSheets[ixti];
                    const book = xti && ctx.supbooks && ctx.supbooks[xti.supbook];
                    let name;
                    if (book && book.type === 'self') {
                        name = ctx.names && ctx.names[idx - 1] && ctx.names[idx - 1].name;
                    } else {
                        name = book && book.externNames && book.externNames[idx - 1];
                        if (name && book.type === 'external') {
                            name = `[${book.name}]!${name}`;
                        }
                    }
                    if (!name) {
                        throw new PTGError(`unknown external name ${ixti}/${idx}`);
                    }
                    stack.push(name);
                    break;
                }
                case 0x1A: { // tRef3d: ixti + RgceLoc (RgceLocRel inside shared formulas)
                    need(6);
                    const prefix = sheetPrefix(readUInt16LE(tokens, pos), ctx);
                    const row = readUInt16LE(tokens, pos + 2);
                    const col = readUInt16LE(tokens, pos + 4);
                    pos += 6;
                    const ref = cellContext ? formatLocRel(row, col, cellContext) : formatLoc(row, col);
                    stack.push(prefix === null ? '#REF!' : prefix + ref);
                    break;
                }
                case 0x1B: { // tArea3d
                    need(10);
                    const prefix = sheetPrefix(readUInt16LE(tokens, pos), ctx);
                    const r1 = readUInt16LE(tokens, pos + 2);
                    const r2 = readUInt16LE(tokens, pos + 4);
                    const c1 = readUInt16LE(tokens, pos + 6);
                    const c2 = readUInt16LE(tokens, pos + 8);
                    pos += 10;
                    const f = cellContext ? (r, c) => formatLocRel(r, c, cellContext) : formatLoc;
                    stack.push(prefix === null ? '#REF!' : `${prefix}${f(r1, c1)}:${f(r2, c2)}`);
                    break;
                }
                case 0x1C: { // tRefErr3d
                    need(6);
                    const prefix = sheetPrefix(readUInt16LE(tokens, pos), ctx);
                    pos += 6;
                    stack.push((prefix || '') + '#REF!');
                    break;
                }
                case 0x1D: { // tAreaErr3d
                    need(10);
                    const prefix = sheetPrefix(readUInt16LE(tokens, pos), ctx);
                    pos += 10;
                    stack.push((prefix || '') + '#REF!');
                    break;
                }
                default:
                    throw new PTGError(`unsupported token 0x${rawToken.toString(16)}`);
            }
            continue;
        }

        switch (rawToken) {
            case 0x01: // tExp: shared / array formula pointer, resolved by the caller
                throw new PTGError('unresolved shared or array formula (tExp)');
            case 0x02: // tTbl: data table (TABLE record), not a formula a user typed
                throw new PTGError('data table (tTbl)');
            case 0x03: binary('+'); break;
            case 0x04: binary('-'); break;
            case 0x05: binary('*'); break;
            case 0x06: binary('/'); break;
            case 0x07: binary('^'); break;
            case 0x08: binary('&'); break;
            case 0x09: binary('<'); break;
            case 0x0A: binary('<='); break;
            case 0x0B: binary('='); break;
            case 0x0C: binary('>='); break;
            case 0x0D: binary('>'); break;
            case 0x0E: binary('<>'); break;
            case 0x0F: binary(' '); break; // intersection
            case 0x10: binary(','); break; // union
            case 0x11: binary(':'); break; // range
            case 0x12: stack.push(`+${pop()}`); break;
            case 0x13: stack.push(`-${pop()}`); break;
            case 0x14: stack.push(`${pop()}%`); break;
            case 0x15: stack.push(`(${pop()})`); break;
            case 0x16: stack.push(''); break; // missing argument
            case 0x17: { // tStr: ShortXLUnicodeString (cch, flags, chars)
                need(2);
                const strLen = tokens[pos];
                const isUnicode = tokens[pos + 1] & 0x01;
                pos += 2;
                need(strLen * (isUnicode ? 2 : 1));
                let str = '';
                for (let i = 0; i < strLen; i++) {
                    str += String.fromCharCode(isUnicode ? readUInt16LE(tokens, pos) : tokens[pos]);
                    pos += isUnicode ? 2 : 1;
                }
                stack.push(`"${str.replace(/"/g, '""')}"`);
                break;
            }
            case 0x19: { // tAttr
                need(3);
                const type = tokens[pos];
                const data = readUInt16LE(tokens, pos + 1);
                pos += 3;
                if (type & 0x04) {
                    // tAttrChoose: a jump table of data + 1 offsets follows
                    need(2 * (data + 1));
                    pos += 2 * (data + 1);
                } else if (type & 0x10) {
                    // tAttrSum: SUM with a single argument
                    stack.push(`SUM(${pop()})`);
                }
                // tAttrVolatile / If / Goto / Space / SpaceVolatile / Baxcel: no effect on the text
                break;
            }
            case 0x1C: // tErr
                need(1);
                stack.push(ERRORS[tokens[pos]] || '#N/A');
                pos++;
                break;
            case 0x1D: // tBool
                need(1);
                stack.push(tokens[pos] ? 'TRUE' : 'FALSE');
                pos++;
                break;
            case 0x1E: // tInt
                need(2);
                stack.push(readUInt16LE(tokens, pos).toString());
                pos += 2;
                break;
            case 0x1F: // tNum
                need(8);
                stack.push(readFloat64LE(tokens, pos).toString());
                pos += 8;
                break;
            default:
                throw new PTGError(`unsupported token 0x${rawToken.toString(16)}`);
        }
    }

    if (stack.length !== 1) {
        throw new PTGError(`malformed formula: ${stack.length} operands left`);
    }
    return stack[0];
}

export { BUILTIN_NAMES };
