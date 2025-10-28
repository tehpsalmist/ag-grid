import type {
    AgColumn,
    AgColumnGroup,
    ExcelExportMultipleSheetParams,
    ExcelExportParams,
    ExcelFactoryMode,
    ExcelRow,
    ExcelStyle,
    IExcelCreator,
    NamedBean,
} from 'ag-grid-community';
import {
    BaseCreator,
    _addGridCommonParams,
    _downloadFile,
    _getHeaderClassesFromColDef,
    _warn,
} from 'ag-grid-community';

import type { ExcelGridSerializingParams, StyleLinkerInterface } from './excelSerializingSession';
import { ExcelSerializingSession } from './excelSerializingSession';
import {
    XLSX_IMAGES,
    XLSX_WORKSHEET_DATA_TABLES,
    XLSX_WORKSHEET_HEADER_FOOTER_IMAGES,
    XLSX_WORKSHEET_IMAGES,
    createXlsxContentTypes,
    createXlsxCore,
    createXlsxDrawing,
    createXlsxDrawingRel,
    createXlsxRelationships,
    createXlsxRels,
    createXlsxSharedStrings,
    createXlsxStylesheet,
    createXlsxTable,
    createXlsxTheme,
    createXlsxVmlDrawing,
    createXlsxVmlDrawingRel,
    createXlsxWorkbook,
    createXlsxWorkbookRels,
    getXlsxFactoryMode,
    resetXlsxFactory,
    setXlsxFactoryMode,
} from './excelXlsxFactory';
import { _normaliseImageExtension } from './files/ooxml/contentTypes';
import { ZipContainer } from './zipContainer/zipContainer';

const createExcelXMLCoreFolderStructure = (zipContainer: ZipContainer): void => {
    zipContainer.addFolders(['_rels/', 'docProps/', 'xl/', 'xl/theme/', 'xl/_rels/', 'xl/worksheets/']);

    if (!XLSX_IMAGES.size) {
        return;
    }

    zipContainer.addFolders(['xl/worksheets/_rels', 'xl/drawings/', 'xl/drawings/_rels', 'xl/media/']);

    let imgCounter = 0;

    XLSX_IMAGES.forEach((value) => {
        const firstImage = value[0].image[0];
        const { base64, imageType } = firstImage;

        zipContainer.addFile(`xl/media/image${++imgCounter}.${_normaliseImageExtension(imageType)}`, base64, true);
    });
};

const createExcelXmlWorksheets = (zipContainer: ZipContainer, data: string[]): void => {
    let imageRelationCounter = 0;
    let headerFooterImageCounter = 0;

    for (let i = 0; i < data.length; i++) {
        const value = data[i];
        zipContainer.addFile(`xl/worksheets/sheet${i + 1}.xml`, value, false);

        const hasImages = XLSX_IMAGES.size > 0 && XLSX_WORKSHEET_IMAGES.has(i);
        const tableData = XLSX_WORKSHEET_DATA_TABLES.size > 0 && XLSX_WORKSHEET_DATA_TABLES.get(i);
        const hasHeaderFooterImages = XLSX_IMAGES.size && XLSX_WORKSHEET_HEADER_FOOTER_IMAGES.has(i);

        if (!hasImages && !tableData && !hasHeaderFooterImages) {
            continue;
        }

        let tableName: string | undefined;
        let drawingIndex: number | undefined;
        let vmlDrawingIndex: number | undefined;

        if (hasImages) {
            createExcelXmlDrawings(zipContainer, i, imageRelationCounter);
            drawingIndex = imageRelationCounter;
            imageRelationCounter++;
        }

        if (hasHeaderFooterImages) {
            createExcelVmlDrawings(zipContainer, i, headerFooterImageCounter);
            vmlDrawingIndex = headerFooterImageCounter;
            headerFooterImageCounter++;
        }

        if (tableData) {
            tableName = tableData.name;
        }

        const worksheetRelFile = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;

        zipContainer.addFile(
            worksheetRelFile,
            createXlsxRelationships({
                tableName,
                drawingIndex,
                vmlDrawingIndex,
            })
        );
    }
};

const createExcelXmlDrawings = (zipContainer: ZipContainer, sheetIndex: number, drawingIndex: number): void => {
    const drawingFolder = 'xl/drawings';
    const drawingFileName = `${drawingFolder}/drawing${drawingIndex + 1}.xml`;
    const relFileName = `${drawingFolder}/_rels/drawing${drawingIndex + 1}.xml.rels`;

    zipContainer.addFile(relFileName, createXlsxDrawingRel(sheetIndex));
    zipContainer.addFile(drawingFileName, createXlsxDrawing(sheetIndex));
};

const createExcelVmlDrawings = (zipContainer: ZipContainer, sheetIndex: number, drawingIndex: number): void => {
    const drawingFolder = 'xl/drawings';
    const drawingFileName = `${drawingFolder}/vmlDrawing${drawingIndex + 1}.vml`;
    const relFileName = `${drawingFolder}/_rels/vmlDrawing${drawingIndex + 1}.vml.rels`;

    zipContainer.addFile(drawingFileName, createXlsxVmlDrawing(sheetIndex));
    zipContainer.addFile(relFileName, createXlsxVmlDrawingRel(sheetIndex));
};

const createExcelXmlTables = (zipContainer: ZipContainer): void => {
    const tablesDataByWorksheet = XLSX_WORKSHEET_DATA_TABLES;
    const worksheetKeys = Array.from(tablesDataByWorksheet.keys());

    for (let i = 0; i < worksheetKeys.length; i++) {
        const sheetIndex = worksheetKeys[i];
        const table = tablesDataByWorksheet.get(sheetIndex);

        if (!table) {
            continue;
        }

        zipContainer.addFile(`xl/tables/${table.name}.xml`, createXlsxTable(table, i));
    }
};

const createExcelXmlCoreSheets = (
    zipContainer: ZipContainer,
    fontSize: number,
    author: string,
    sheetLen: number,
    activeTab: number
): void => {
    zipContainer.addFile('xl/workbook.xml', createXlsxWorkbook(activeTab));
    zipContainer.addFile('xl/styles.xml', createXlsxStylesheet(fontSize));
    zipContainer.addFile('xl/sharedStrings.xml', createXlsxSharedStrings());
    zipContainer.addFile('xl/theme/theme1.xml', createXlsxTheme());
    zipContainer.addFile('xl/_rels/workbook.xml.rels', createXlsxWorkbookRels(sheetLen));
    zipContainer.addFile('docProps/core.xml', createXlsxCore(author));
    zipContainer.addFile('[Content_Types].xml', createXlsxContentTypes(sheetLen));
    zipContainer.addFile('_rels/.rels', createXlsxRels());
};

const createExcelFileForExcel = (
    zipContainer: ZipContainer,
    data: string[],
    options: {
        columns?: string[];
        rowCount?: number;
        fontSize?: number;
        author?: string;
        activeTab?: number;
    } = {}
): boolean => {
    if (!data || data.length === 0) {
        _warn(159);
        resetXlsxFactory();
        return false;
    }

    const { fontSize = 11, author = 'AG Grid', activeTab = 0 } = options;

    const len = data.length;
    const activeTabWithinBounds = Math.max(Math.min(activeTab, len - 1), 0);

    createExcelXMLCoreFolderStructure(zipContainer);
    createExcelXmlTables(zipContainer);
    createExcelXmlWorksheets(zipContainer, data);
    createExcelXmlCoreSheets(zipContainer, fontSize, author, len, activeTabWithinBounds);

    resetXlsxFactory();

    return true;
};

const getMultipleSheetsAsExcelCompressed = (params: ExcelExportMultipleSheetParams): Promise<Blob | undefined> => {
    const { data, fontSize, author, activeSheetIndex } = params;
    const mimeType = params.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const zipContainer = new ZipContainer();

    if (
        !createExcelFileForExcel(zipContainer, data, {
            author,
            fontSize,
            activeTab: activeSheetIndex,
        })
    ) {
        return Promise.resolve(undefined);
    }

    const zipFile = zipContainer.getZipFile(mimeType);

    if (params.password) {
        return zipFile.then((file) => {
            // encrypt the zip file with the provided password
            return file;
        });
    }

    return zipFile;
};

const ENCRYPTION_INFO_PREFIX = Uint8Array.from([0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00]); // First 4 bytes are the version number, second 4 bytes are reserved.
const PACKAGE_ENCRYPTION_CHUNK_SIZE = 4096;
const PACKAGE_OFFSET = 8; // First 8 bytes are the size of the stream

// Block keys used for encryption
const BLOCK_KEYS = {
    dataIntegrity: {
        hmacKey: Uint8Array.from([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]),
        hmacValue: Uint8Array.from([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]),
    },
    key: Uint8Array.from([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]),
    verifierHash: {
        input: Uint8Array.from([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]),
        value: Uint8Array.from([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]),
    },
};

const encryptBlobWithPassword = async (blob: Blob, password: string): Promise<Blob> => {
    // Generate a random key to use to encrypt the document. Excel uses 32 bytes. We'll use the password to encrypt this key.
    // N.B. The number of bits needs to correspond to an algorithm available in crypto (e.g. aes-256-cbc).
    const blobAsBuffer = await blob.arrayBuffer();
    const packageKey = crypto.getRandomValues(new Uint8Array(32));
    const packageSalt = crypto.getRandomValues(new Uint8Array(16));
    const keySalt = crypto.getRandomValues(new Uint8Array(16));

    // Create the encryption info. We'll use this for all of the encryption operations and for building the encryption info XML entry
    const encryptionInfo: Record<
        string,
        | {
              cipherAlgorithm: string;
              cipherChaining: string;
              saltValue: Uint8Array;
              hashAlgorithm: string;
              hashSize: number;
              blockSize: number;
              keyBits: number;
              spinCount?: number;
          }
        | {
              encryptedHmacKey: string;
              encryptedHmacValue: string;
          }
    > = {
        package: {
            // Info on the encryption of the package.
            cipherAlgorithm: 'AES', // Cipher algorithm to use. Excel uses AES.
            cipherChaining: 'ChainingModeCBC', // Cipher chaining mode to use. Excel uses CBC.
            saltValue: packageSalt, // Random value to use as encryption salt. Excel uses 16 bytes.
            hashAlgorithm: 'SHA512', // Hash algorithm to use. Excel uses SHA512.
            hashSize: 64, // The size of the hash in bytes. SHA512 results in 64-byte hashes
            blockSize: 16, // The number of bytes used to encrypt one block of data. It MUST be at least 2, no greater than 4096, and a multiple of 2. Excel uses 16
            keyBits: packageKey.length * 8, // The number of bits in the package key.
        },
        key: {
            // Info on the encryption of the package key.
            cipherAlgorithm: 'AES', // Cipher algorithm to use. Excel uses AES.
            cipherChaining: 'ChainingModeCBC', // Cipher chaining mode to use. Excel uses CBC.
            saltValue: keySalt, // Random value to use as encryption salt. Excel uses 16 bytes.
            hashAlgorithm: 'SHA512', // Hash algorithm to use. Excel uses SHA512.
            hashSize: 64, // The size of the hash in bytes. SHA512 results in 64-byte hashes
            blockSize: 16, // The number of bytes used to encrypt one block of data. It MUST be at least 2, no greater than 4096, and a multiple of 2. Excel uses 16
            spinCount: 100000, // The number of times to iterate on a hash of a password. It MUST NOT be greater than 10,000,000. Excel uses 100,000.
            keyBits: 256, // The length of the key to generate from the password. Must be a multiple of 8. Excel uses 256.
        },
    };

    /* Package Encryption */

    // Encrypt package using the package key.
    const encryptedPackage = _cryptPackage(
        true,
        encryptionInfo.package.cipherAlgorithm,
        encryptionInfo.package.cipherChaining,
        encryptionInfo.package.hashAlgorithm,
        encryptionInfo.package.blockSize,
        encryptionInfo.package.saltValue,
        packageKey,
        blobAsBuffer
    );

    /* Data Integrity */

    // Create the data integrity fields used by clients for integrity checks.
    // First generate a random array of bytes to use in HMAC. The docs say to use the same length as the key salt, but Excel seems to use 64.
    const hmacKey = crypto.getRandomValues(new Uint8Array(64));

    // Then create an initialization vector using the package encryption info and the appropriate block key.
    const hmacKeyIV = _createIV(
        encryptionInfo.package.hashAlgorithm,
        encryptionInfo.package.saltValue,
        encryptionInfo.package.blockSize,
        BLOCK_KEYS.dataIntegrity.hmacKey
    );

    // Use the package key and the IV to encrypt the HMAC key
    const encryptedHmacKey = _crypt(
        true,
        encryptionInfo.package.cipherAlgorithm,
        encryptionInfo.package.cipherChaining,
        packageKey,
        hmacKeyIV,
        hmacKey
    );

    // Now create the HMAC
    const hmacValue = _hmac(encryptionInfo.package.hashAlgorithm, hmacKey, encryptedPackage);

    // Next generate an initialization vector for encrypting the resulting HMAC value.
    const hmacValueIV = _createIV(
        encryptionInfo.package.hashAlgorithm,
        encryptionInfo.package.saltValue,
        encryptionInfo.package.blockSize,
        BLOCK_KEYS.dataIntegrity.hmacValue
    );

    // Now encrypt the value
    const encryptedHmacValue = _crypt(
        true,
        encryptionInfo.package.cipherAlgorithm,
        encryptionInfo.package.cipherChaining,
        packageKey,
        hmacValueIV,
        hmacValue
    );

    // Put the encrypted key and value on the encryption info
    encryptionInfo.dataIntegrity = {
        encryptedHmacKey,
        encryptedHmacValue,
    };

    /* Key Encryption */

    // Convert the password to an encryption key
    const key = _convertPasswordToKey(
        password,
        encryptionInfo.key.hashAlgorithm,
        encryptionInfo.key.saltValue,
        encryptionInfo.key.spinCount,
        encryptionInfo.key.keyBits,
        BLOCK_KEYS.key
    );

    // Encrypt the package key with the
    encryptionInfo.key.encryptedKeyValue = _crypt(
        true,
        encryptionInfo.key.cipherAlgorithm,
        encryptionInfo.key.cipherChaining,
        key,
        encryptionInfo.key.saltValue,
        packageKey
    );

    /* Verifier hash */

    // Create a random byte array for hashing
    const verifierHashInput = crypto.getRandomValues(new Uint8Array(16));

    // Create an encryption key from the password for the input
    const verifierHashInputKey = _convertPasswordToKey(
        password,
        encryptionInfo.key.hashAlgorithm,
        encryptionInfo.key.saltValue,
        encryptionInfo.key.spinCount,
        encryptionInfo.key.keyBits,
        BLOCK_KEYS.verifierHash.input
    );

    // Use the key to encrypt the verifier input
    encryptionInfo.key.encryptedVerifierHashInput = _crypt(
        true,
        encryptionInfo.key.cipherAlgorithm,
        encryptionInfo.key.cipherChaining,
        verifierHashInputKey,
        encryptionInfo.key.saltValue,
        verifierHashInput
    );

    // Create a hash of the input
    const verifierHashValue = _hash(encryptionInfo.key.hashAlgorithm, verifierHashInput);

    // Create an encryption key from the password for the hash
    const verifierHashValueKey = _convertPasswordToKey(
        password,
        encryptionInfo.key.hashAlgorithm,
        encryptionInfo.key.saltValue,
        encryptionInfo.key.spinCount,
        encryptionInfo.key.keyBits,
        BLOCK_KEYS.verifierHash.value
    );

    // Use the key to encrypt the hash value
    encryptionInfo.key.encryptedVerifierHashValue = _crypt(
        true,
        encryptionInfo.key.cipherAlgorithm,
        encryptionInfo.key.cipherChaining,
        verifierHashValueKey,
        encryptionInfo.key.saltValue,
        verifierHashValue
    );

    // Build the encryption info buffer
    const encryptionInfoBuffer = _buildEncryptionInfo(encryptionInfo);

    // Create a new CFB
    let output = cfb.utils.cfb_new();

    // Add the encryption info and encrypted package
    cfb.utils.cfb_add(output, 'EncryptionInfo', encryptionInfoBuffer);
    cfb.utils.cfb_add(output, 'EncryptedPackage', encryptedPackage);

    // Delete the SheetJS entry that is added at initialization
    cfb.utils.cfb_del(output, '\u0001Sh33tJ5');

    // Write to a buffer and return
    output = cfb.write(output);

    return output;
};

/**
 * Encrypt/decrypt the package
 * @param encrypt - True to encrypt, false to decrypt
 * @param cipherAlgorithm - The cipher algorithm
 * @param cipherChaining - The cipher chaining mode
 * @param hashAlgorithm - The hash algorithm
 * @param blockSize - The IV block size
 * @param saltValue - The salt
 * @param key - The encryption key
 * @param input - The package input
 * @returns The output
 * @private
 */
const _cryptPackage = (
    encrypt: boolean,
    cipherAlgorithm: string,
    cipherChaining: string,
    hashAlgorithm: string,
    blockSize: number,
    saltValue: ArrayBuffer,
    key: ArrayBuffer,
    input: ArrayBuffer
) => {
    // The first 8 bytes is supposed to be the length, but it seems like it is really the length - 4..
    const outputChunks = [];
    const offset = encrypt ? 0 : PACKAGE_OFFSET;

    // The package is encoded in chunks. Encrypt/decrypt each and concat.
    let i = 0,
        start = 0,
        end = 0;
    while (end < input.length) {
        start = end;
        end = start + PACKAGE_ENCRYPTION_CHUNK_SIZE;
        if (end > input.length) end = input.length;

        // Grab the next chunk
        let inputChunk = input.slice(start + offset, end + offset);

        // Pad the chunk if it is not an integer multiple of the block size
        const remainder = inputChunk.length % blockSize;
        if (remainder) inputChunk = Buffer.concat([inputChunk, Buffer.alloc(blockSize - remainder)]);

        // Create the initialization vector
        const iv = this._createIV(hashAlgorithm, saltValue, blockSize, i);

        // Encrypt/decrypt the chunk and add it to the array
        const outputChunk = this._crypt(encrypt, cipherAlgorithm, cipherChaining, key, iv, inputChunk);
        outputChunks.push(outputChunk);

        i++;
    }

    // Concat all of the output chunks.
    let output = Buffer.concat(outputChunks);

    if (encrypt) {
        // Put the length of the package in the first 8 bytes
        output = Buffer.concat([this._createUInt32LEBuffer(input.length, PACKAGE_OFFSET), output]);
    } else {
        // Truncate the buffer to the size in the prefix
        const length = input.readUInt32LE(0);
        output = output.slice(0, length);
    }

    return output;
};
export const getMultipleSheetsAsExcel = (params: ExcelExportMultipleSheetParams): Blob | undefined => {
    const { data, fontSize, author, activeSheetIndex } = params;
    const mimeType = params.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const zipContainer = new ZipContainer();

    if (
        !createExcelFileForExcel(zipContainer, data, {
            author,
            fontSize,
            activeTab: activeSheetIndex,
        })
    ) {
        return;
    }

    return zipContainer.getUncompressedZipFile(mimeType);
};

export const exportMultipleSheetsAsExcel = (params: ExcelExportMultipleSheetParams) => {
    const { fileName = 'export.xlsx' } = params;

    getMultipleSheetsAsExcelCompressed(params).then((contents) => {
        if (contents) {
            const downloadFileName = typeof fileName === 'function' ? fileName() : fileName;

            _downloadFile(downloadFileName, contents);
        }
    });
};

export class ExcelCreator
    extends BaseCreator<ExcelRow[], ExcelSerializingSession, ExcelExportParams>
    implements NamedBean, IExcelCreator
{
    beanName = 'excelCreator' as const;

    protected getMergedParams(params?: ExcelExportParams): ExcelExportParams {
        const baseParams = this.gos.get('defaultExcelExportParams');
        return Object.assign({}, baseParams, params);
    }

    protected export(userParams?: ExcelExportParams): void {
        if (this.isExportSuppressed()) {
            _warn(160);
            return;
        }

        const mergedParams = this.getMergedParams(userParams);
        const data = this.getData(mergedParams);

        const exportParams: ExcelExportMultipleSheetParams = {
            data: [data],
            fontSize: mergedParams.fontSize,
            author: mergedParams.author,
            mimeType: mergedParams.mimeType,
        };
        if (mergedParams.password) {
            exportParams.password = mergedParams.password;
        }

        this.packageCompressedFile(exportParams).then((packageFile) => {
            if (packageFile) {
                const { fileName } = mergedParams;
                const providedFileName =
                    typeof fileName === 'function' ? fileName(_addGridCommonParams(this.gos, {})) : fileName;

                _downloadFile(this.getFileName(providedFileName), packageFile);
            }
        });
    }

    public exportDataAsExcel(params?: ExcelExportParams): void {
        this.export(params);
    }

    public getDataAsExcel(params?: ExcelExportParams): Blob | string | undefined {
        const mergedParams = this.getMergedParams(params);
        const data = this.getData(mergedParams);

        const exportParams: ExcelExportMultipleSheetParams = {
            data: [data],
            fontSize: mergedParams.fontSize,
            author: mergedParams.author,
            mimeType: mergedParams.mimeType,
        };

        return this.packageFile(exportParams);
    }

    public setFactoryMode(factoryMode: ExcelFactoryMode): void {
        setXlsxFactoryMode(factoryMode);
    }

    public getFactoryMode(): ExcelFactoryMode {
        return getXlsxFactoryMode();
    }

    public getSheetDataForExcel(params: ExcelExportParams): string {
        const mergedParams = this.getMergedParams(params);
        return this.getData(mergedParams);
    }

    public getMultipleSheetsAsExcel(params: ExcelExportMultipleSheetParams): Blob | undefined {
        return getMultipleSheetsAsExcel(params);
    }

    public exportMultipleSheetsAsExcel(params: ExcelExportMultipleSheetParams): void {
        exportMultipleSheetsAsExcel(params);
    }

    public getDefaultFileExtension(): 'xlsx' {
        return 'xlsx';
    }

    public createSerializingSession(params: ExcelExportParams): ExcelSerializingSession {
        const { colModel, colNames, rowGroupColsSvc, valueSvc, gos } = this.beans;

        const config: ExcelGridSerializingParams = {
            ...params,
            colModel,
            colNames,
            rowGroupColsSvc,
            valueSvc,
            gos,
            suppressRowOutline: params.suppressRowOutline || params.skipRowGroups,
            headerRowHeight: params.headerRowHeight || params.rowHeight,
            baseExcelStyles: gos.get('excelStyles') || [],
            rightToLeft: params.rightToLeft ?? gos.get('enableRtl'),
            styleLinker: this.styleLinker.bind(this),
        };

        return new ExcelSerializingSession(config);
    }

    private styleLinker(params: StyleLinkerInterface): string[] {
        const { rowType, rowIndex, value, column, columnGroup, node } = params;
        const isHeader = rowType === 'HEADER';
        const isGroupHeader = rowType === 'HEADER_GROUPING';
        const col = (isHeader ? column : columnGroup) as AgColumn | AgColumnGroup | null;
        let headerClasses: string[] = [];
        const { gos, cellStyles } = this.beans;

        if (isHeader || isGroupHeader) {
            headerClasses.push('header');
            if (isGroupHeader) {
                headerClasses.push('headerGroup');
            }

            if (col) {
                headerClasses = headerClasses.concat(
                    _getHeaderClassesFromColDef(
                        col.getDefinition(),
                        gos,
                        (column as AgColumn) || null,
                        (columnGroup as AgColumnGroup) || null
                    )
                );
            }

            return headerClasses;
        }

        const styles = gos.get('excelStyles');

        const applicableStyles: string[] = ['cell'];

        if (!styles?.length) {
            return applicableStyles;
        }

        const styleIds: string[] = styles.map((it: ExcelStyle) => {
            return it.id;
        });

        const colDef = (column as AgColumn).getDefinition();
        cellStyles?.processAllCellClasses(
            colDef,
            _addGridCommonParams(gos, {
                value,
                data: node!.data,
                node: node!,
                colDef,
                column: column!,
                rowIndex: rowIndex,
            }),
            (className: string) => {
                if (styleIds.indexOf(className) > -1) {
                    applicableStyles.push(className);
                }
            }
        );

        return applicableStyles.sort((left: string, right: string): number => {
            return styleIds.indexOf(left) < styleIds.indexOf(right) ? -1 : 1;
        });
    }

    public isExportSuppressed(): boolean {
        return this.gos.get('suppressExcelExport');
    }

    private packageCompressedFile(params: ExcelExportMultipleSheetParams): Promise<Blob | undefined> {
        return getMultipleSheetsAsExcelCompressed(params);
    }

    private packageFile(params: ExcelExportMultipleSheetParams): Blob | undefined {
        return getMultipleSheetsAsExcel(params);
    }
}
