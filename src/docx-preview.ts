import { WordDocument } from './word-document';

import { DocumentParser } from './document-parser';

// HTML Render Asynchronously
import { HtmlRenderer } from './html-renderer';

// HTML Render Synchronously
import { HtmlRendererSync } from "./html-renderer-sync";

export interface Options {
	breakPages: boolean;                    //enables page breaking on page breaks
	className: string;                      //class name/prefix for default and document style classes

	ignoreFonts: boolean;                   //disables fonts rendering
	ignoreHeight: boolean;                  //disables rendering height of page
	ignoreImageWrap: boolean;               //disables image text wrap setting
	ignoreLastRenderedPageBreak: boolean;   //disables page breaking on lastRenderedPageBreak elements
	ignoreTableWrap: boolean;               //disables table's text wrap setting
	ignoreWidth: boolean;                   //disables rendering width of page

	inWrapper: boolean;                     //enables rendering of wrapper around document content

	renderChanges: boolean;                 //enables experimental rendering of document changes (inserions/deletions)
	renderEndnotes: boolean;                //enables endnotes rendering
	renderFooters: boolean;                 //enables footers rendering
	renderFootnotes: boolean;               //enables footnotes rendering
	renderHeaders: boolean;                 //enables headers rendering

	trimXmlDeclaration: boolean;            //if true, xml declaration will be removed from xml documents before parsing
	useBase64URL: boolean;                  //if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used

	debug: boolean;                         //enables additional logging
	experimental: boolean;                  //enables experimental features (tab stops calculation)
}

export const defaultOptions: Options = {
	breakPages: true,
	className: "docx",

	ignoreFonts: false,
	ignoreHeight: false,
	ignoreImageWrap: false,
	ignoreLastRenderedPageBreak: true,
	ignoreTableWrap: true,
	ignoreWidth: false,

	inWrapper: true,

	renderChanges: false,
	renderEndnotes: true,
	renderFooters: true,
	renderFootnotes: true,
	renderHeaders: true,

	trimXmlDeclaration: true,
	useBase64URL: true,

	debug: false,
	experimental: false,
}

// Document Parser
export function parseAsync(data: Blob | any, userOptions: Partial<Options> = null): Promise<any> {
	// assign defaultOptions
	const ops = { ...defaultOptions, ...userOptions };
	// 加载blob对象，根据DocumentParser转换规则，blob对象 => Object对象
	return WordDocument.load(data, new DocumentParser(ops), ops);
}

// Document Render
export async function renderDocument(document: any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, sync: boolean = true, userOptions?: Partial<Options>): Promise<any> {
	// assign defaultOptions
	const ops = { ...defaultOptions, ...userOptions };
	// HTML渲染器实例
	const renderer = sync ? new HtmlRendererSync() : new HtmlRenderer();
	// Asociar el renderer con el documento para limpieza posterior
	if (document && typeof document.setRenderer === 'function') {
		document.setRenderer(renderer);
	}
	// Object对象 => HTML标签
	await renderer.render(document, bodyContainer, styleContainer, ops);
}

// Mapa para trackear documentos asociados con contenedores para auto-dispose
const containerDocumentMap = new WeakMap<HTMLElement, any>();

// Mapa para prevenir renderizados concurrentes en el mismo contenedor
const containerRenderingLock = new WeakMap<HTMLElement, Promise<any>>();

// Render Synchronously
export async function renderSync(data: Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, userOptions: Partial<Options> = null): Promise<any> {
	// Verificar si ya hay un renderizado en progreso para este contenedor
	const currentRendering = containerRenderingLock.get(bodyContainer);
	if (currentRendering) {
		// Esperar a que termine el renderizado actual antes de continuar
		await currentRendering;
	}

	// Crear una promesa para este renderizado y guardarla en el lock
	const renderingPromise = (async () => {
		try {
			// Limpiar contenedores antes de renderizar para evitar fugas de memoria
			bodyContainer.innerHTML = '';
			if (styleContainer) {
				styleContainer.innerHTML = '';
			}

			// Llamar dispose en documento anterior si existe
			const previousDoc = containerDocumentMap.get(bodyContainer);
			if (previousDoc && typeof previousDoc.dispose === 'function') {
				previousDoc.dispose();
			}

			// parse document data
			const doc = await parseAsync(data, userOptions);
			// render document
			await renderDocument(doc, bodyContainer, styleContainer, true, userOptions);

			// Trackear el documento actual para auto-dispose en futuras renderizaciones
			containerDocumentMap.set(bodyContainer, doc);

			return doc;
		} finally {
			// Remover el lock cuando termine
			containerRenderingLock.delete(bodyContainer);
		}
	})();

	// Guardar la promesa en el lock
	containerRenderingLock.set(bodyContainer, renderingPromise);

	return renderingPromise;
}

// Render Asynchronously
export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
	// Verificar si ya hay un renderizado en progreso para este contenedor
	const currentRendering = containerRenderingLock.get(bodyContainer);
	if (currentRendering) {
		// Esperar a que termine el renderizado actual antes de continuar
		await currentRendering;
	}

	// Crear una promesa para este renderizado y guardarla en el lock
	const renderingPromise = (async () => {
		try {
			// Limpiar contenedores antes de renderizar para evitar fugas de memoria
			bodyContainer.innerHTML = '';
			if (styleContainer) {
				styleContainer.innerHTML = '';
			}

			// Llamar dispose en documento anterior si existe
			const previousDoc = containerDocumentMap.get(bodyContainer);
			if (previousDoc && typeof previousDoc.dispose === 'function') {
				previousDoc.dispose();
			}

			const doc = await parseAsync(data, userOptions);
			await renderDocument(doc, bodyContainer, styleContainer, false, userOptions);

			// Trackear el documento actual para auto-dispose en futuras renderizaciones
			containerDocumentMap.set(bodyContainer, doc);

			return doc;
		} finally {
			// Remover el lock cuando termine
			containerRenderingLock.delete(bodyContainer);
		}
	})();

	// Guardar la promesa en el lock
	containerRenderingLock.set(bodyContainer, renderingPromise);

	return renderingPromise;
}
