# Third-party notices

## Shipped code

No third-party code ships with this extension. The compound file reader, the ZIP reader (it uses the zlib module of Node.js), the VBA project reader and the VBA decompression in `src/core` are written for this extension from the public Microsoft Open Specifications [MS-CFB], [MS-OVBA], [MS-PPT] and [MS-XLS]/[MS-XLSB] documents. Every other source file and the icon are written for this extension (MIT, see LICENSE).

## Samples

`samples/sample-macros.xlsm` was made for this extension by `scripts/make-samples.js`. Its names and amounts are made up, and its code is harmless.

## Test files (in the source repository, not in the extension package)

The unit tests read some real Office files. They are not part of the packaged extension.

- `test/fixtures/real/SimpleMacro.xls`, `SimpleMacro.xlsm`, `SimpleMacro.doc`, `SimpleMacro.docm`, `SimpleMacro.ppt`, `SimpleMacro.pptm`, `xlmmacro.xlsm`, `Simple.xlsb`, `SimpleMacro-xls.vba` and `SimpleMacro-ppt.vba` are test files of Apache POI (https://github.com/apache/poi, folder test-data), licensed under the Apache License 2.0. Copyright The Apache Software Foundation.
- `test/fixtures/real/VBA-JSON-Specs.xlsm` and `JsonConverter.bas` are from VBA-JSON by Tim Hall (https://github.com/VBA-tools/VBA-JSON), licensed under the MIT License.
- The other files in `test/fixtures` were made for this extension by `scripts/make-samples.js`.
