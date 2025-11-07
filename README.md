# docx-preview-sync
The docx document **synchronously** rendering library with dynamic document injection support

[![npm version](https://badge.fury.io/js/docx-preview-sync.svg)](https://www.npmjs.com/package/docx-preview-sync)

## Credits
This library is inspired by the [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) library. Thanks to the original author [VolodymyrBaydalka](https://github.com/VolodymyrBaydalka) for the foundational work.

## Introduction
This is an enhanced fork that maintains the original synchronous rendering capabilities while adding critical improvements:

### Key Enhancements
- **Memory Leak Fixes**: Fixed critical memory issues that caused RAM to grow significantly during rapid document reloads. The original implementation had problems with proper cleanup of Konva stages, DOM references, and ObjectURLs.
- **Intelligent Throttling**: Automatic rendering throttle system (200ms minimum spacing) prevents memory bloat from rapid successive render calls while maintaining responsive document updates.
- **Dynamic Document Injection**: Ability to inject and replace parsed documents in the current renderer instance without full re-initialization, enabling efficient document switching in applications.
- **Stable Resource Management**: Improved garbage collection handling with WeakMaps for document tracking and comprehensive disposal chains.

### Original Features (Maintained)
The library continues to support synchronous rendering with page breaking capabilities, allowing detection and splitting of HTML elements across pages.

**Note**: This library prioritizes correctness and resource efficiency over raw performance. Synchronous rendering is inherently slower than asynchronous alternatives.

Online Demo
-----
[https://millet0328.github.io/docx-preview-sync/](https://millet0328.github.io/docx-preview-sync/)

## Usage

### Basic Synchronous Rendering

#### Package managers
Install library in your Node.js powered apps with the npm package:

```shell
npm install docx-preview-sync
```

```typescript
import { renderSync } from 'docx-preview-sync';

// fetch document Blob, maybe from input with type = file
let docData: Blob = document.querySelector('input').files[0];

// synchronously rendering function
let wordDocument = await renderSync(docData, document.getElementById("container"));

// if you need to get the word document object
console.log("docx document object", wordDocument);
```

### Dynamic Document Injection

One of the key enhancements is the ability to inject a new document into an existing renderer instance without re-initialization:

```typescript
import { renderSync, replaceParsedDocument } from 'docx-preview-sync';

// Initial render
let wordDocument = await renderSync(firstDocBlob, document.getElementById("container"));

// Later, inject a new document into the same container without full re-render
let newDocument = await replaceParsedDocument(
  secondDocBlob,
  document.getElementById("container"),
  wordDocument  // pass the existing document instance
);
```

This is particularly useful for:
- Switching between documents in a single viewer
- Efficient document management in multi-document applications
- Reduced initialization overhead when replacing documents

### Static HTML without a build step

```html
<!--dependencies-->
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/lodash@4.17.21/lodash.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/konva@9.3.6/konva.min.js"></script>

<script src="./js/docx-preview.min.js"></script>

<body>
    ...
    <div id="container"></div>
    ...
</body>

<script>
  // fectch document Blob,maybe from input width type = file
  let docData = document.querySelector('input').files[0];

  // synchronously rendering function
  docx.renderSync(docData, document.getElementById("container"))
      .then(wordDocument => {
        // if you need to get the Word document object
        console.log("docx document object", wordDocument);
      });
</script>
```

API
---
### renderSync

Render HTML5 Elements Synchronously.

renderSync = praseAsync + renderDocument(sync:true)

```typescript
// renders document into specified element synchronously
renderSync(
    document: Blob | ArrayBuffer | Uint8Array, // could be any type that supported by JSZip.loadAsync
    bodyContainer: HTMLElement, //element to render document content,
    styleContainer: HTMLElement, //element to render document styles, numbeings, fonts. If null, bodyContainer will be used.
    options: {
        breakPages: boolean = true, //enables page breaking on page breaks
        className: string = "docx", //class name/prefix for default and document style classes
        ignoreFonts: boolean = false, //disables fonts rendering
        ignoreHeight: boolean = false, //disables rendering height of page
        ignoreImageWrap: boolean = false, //disables image text wrap setting
        ignoreLastRenderedPageBreak: boolean = true, //disables page breaking on lastRenderedPageBreak elements
        ignoreTableWrap: boolean = true, //disables table's text wrap setting
        ignoreWidth: boolean = false, //disables rendering width of page
        inWrapper: boolean = true, //enables rendering of wrapper around document content
        renderChanges: false, //enables experimental rendering of document changes (inserions/deletions)
        renderEndnotes: true, //enables endnotes rendering
        renderFooters: true, //enables footers rendering
        renderFootnotes: true, //enables footnotes rendering
        renderHeaders: true, //enables headers rendering
        trimXmlDeclaration: boolean = true, //if true, xml declaration will be removed from xml documents before parsing
        useBase64URL: boolean = false, //if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used
        debug: boolean = false, //enables additional logging
        experimental: boolean = false, //enables experimental features (tab stops calculation)
    }): Promise<WordDocument>
```

### renderAsync

Render HTML5 Elements Asynchronously

renderSync = praseAsync + renderDocument(sync:false)

```typescript
// renders document into specified element synchronously
renderAsync(
    document: Blob | ArrayBuffer | Uint8Array, // could be any type that supported by JSZip.loadAsync
    bodyContainer: HTMLElement, //element to render document content,
    styleContainer: HTMLElement, //element to render document styles, numbeings, fonts. If null, bodyContainer will be used.
    options: {
        breakPages: boolean = true, //enables page breaking on page breaks
        className: string = "docx", //class name/prefix for default and document style classes
        ignoreFonts: boolean = false, //disables fonts rendering
        ignoreHeight: boolean = false, //disables rendering height of page
        ignoreImageWrap: boolean = false, //disables image text wrap setting
        ignoreLastRenderedPageBreak: boolean = true, //disables page breaking on lastRenderedPageBreak elements
        ignoreTableWrap: boolean = true, //disables table's text wrap setting
        ignoreWidth: boolean = false, //disables rendering width of page
        inWrapper: boolean = true, //enables rendering of wrapper around document content
        renderChanges: false, //enables experimental rendering of document changes (inserions/deletions)
        renderEndnotes: true, //enables endnotes rendering
        renderFooters: true, //enables footers rendering
        renderFootnotes: true, //enables footnotes rendering
        renderHeaders: true, //enables headers rendering
        trimXmlDeclaration: boolean = true, //if true, xml declaration will be removed from xml documents before parsing
        useBase64URL: boolean = false, //if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used
        debug: boolean = false, //enables additional logging
        experimental: boolean = false, //enables experimental features (tab stops calculation)
    }): Promise<WordDocument>
```

### parseAsync

Document Parser Asynchronously

parse document and return internal document object. this could be used to modify document object before rendering

```typescript
// ==== experimental / internal API ====
parseAsync(document: Blob | ArrayBuffer | Uint8Array, options: Options): Promise<WordDocument>
```

### renderDocument

render internal document object into specified container.

sync is a boolean parameter which enables synchronous rendering or asynchronous.

```typescript
// ==== experimental / internal API ====
renderDocument(
  wordDocument: WordDocument, 
  bodyContainer: HTMLElement, 
  styleContainer: HTMLElement,
  sync: boolean,
  options: Options,
): Promise<void>
```

## Goals
* Render/convert DOCX documents into HTML with semantic preservation
* Support efficient page breaking with synchronous rendering
* Provide robust memory management for long-running applications
* Enable dynamic document injection for flexible document switching
* Parse all elements from Office Open XML specification with proper resource cleanup

This library is limited by HTML capabilities. If you need to render documents on canvas, try the OnlyOffice library.

Partially Supported Namespaces
------------------
1. [x] DocumentFormat.OpenXml.Wordprocessing
2. [x] DocumentFormat.OpenXml.Math

Not Supported Namespaces
------------------------
1. [ ] DocumentFormat.OpenXml.Drawing
2. [ ] DocumentFormat.OpenXml.Drawing.Charts
3. [ ] DocumentFormat.OpenXml.InkML
4. [ ] DocumentFormat.OpenXml.Vml

## Page Breaking Behavior

Currently, the library breaks pages:

- When user/manual page breaks are inserted (`<w:br w:type="page"/>`)
- When application page breaks are inserted (`<w:lastRenderedPageBreak/>`) - typically added by MS Word or similar editors (controlled by `ignoreLastRenderedPageBreak` option)
- When page settings for a paragraph are changed (e.g., portrait to landscape)

**Note on Real-time Breaking**: Real-time page breaking is not implemented as it would require continuous size recalculation, significantly impacting performance.

### Recommendations for Optimal Page Breaking

For consistent and reliable page breaking:
- Use manual break points (`<w:br w:type="page"/>`) in your documents where critical
- Allow editors like MS Word to insert `<w:lastRenderedPageBreak/>` markers
- Set `ignoreLastRenderedPageBreak: false` in options to respect these markers

## Memory Management

This fork includes significant memory management improvements:

- **Proper Resource Cleanup**: Comprehensive disposal chains ensure DOM elements, Konva stages, and Object URLs are cleaned up correctly
- **Throttled Rendering**: Automatic 200ms minimum spacing prevents memory growth from rapid successive renders
- **WeakMap Tracking**: Uses WeakMaps for document and rendering state to enable proper garbage collection
- **Garbage Collection Awareness**: Includes GC timing adjustments to allow browser garbage collection to execute

These improvements prevent the memory bloat issues present in the original library during intensive usage scenarios.

## Performance Notes

Due to the synchronous rendering approach:
- Document rendering may be slower than asynchronous alternatives
- Large or complex documents may cause UI blocking during initial render
- The trade-off is precise page breaking capability and robust resource management

For applications requiring ultra-high performance, consider the async version or libraries optimized for performance.

## Links

- **Original Library**: [docx-preview](https://github.com/VolodymyrBaydalka/docxjs)
- **NPM Package**: [docx-preview-sync](https://www.npmjs.com/package/docx-preview-sync)
