import { WordDocument } from './word-document';

import { DocumentParser } from './document-parser';

// HTML Render Asynchronously
import { HtmlRenderer } from './html-renderer';

// HTML Render Synchronously
import { HtmlRendererSync } from "./html-renderer-sync";
import { debugStats } from './debug-stats';

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

	reuseRenderer?: boolean;                //si true permite reutilizar renderer entre reemplazos (experimental)
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

	reuseRenderer: false,
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

// Mapa para implementar throttling (garantizar espera entre renders)
// Almacena { lastRenderTime, pendingTimeout, pendingRequest, pendingResolvers } para cada contenedor
const containerThrottleState = new WeakMap<HTMLElement, {
	lastRenderTime: number;
	pendingTimeout: ReturnType<typeof setTimeout> | null;
	pendingRequest: (() => Promise<any>) | null;
	pendingResolvers: ((result: any) => void)[];
}>();

// Configuración global de throttling
const RENDER_THROTTLE_MS = 200;  // Esperar 200ms entre renders

// Función para aplicar throttling a un render
// Garantiza que NO se ejecuten dos renders en el mismo contenedor más frecuentemente que RENDER_THROTTLE_MS
// IMPORTANTE: Si múltiples solicitudes llegan durante el throttle window, se ejecuta la ÚLTIMA
async function throttledRender<T>(
	bodyContainer: HTMLElement,
	renderFn: () => Promise<T>
): Promise<T> {
	let throttleState = containerThrottleState.get(bodyContainer);
	if (!throttleState) {
		throttleState = {
			lastRenderTime: 0,
			pendingTimeout: null,
			pendingRequest: null,
			pendingResolvers: []
		};
		containerThrottleState.set(bodyContainer, throttleState);
	}

	const now = Date.now();
	const timeSinceLastRender = now - throttleState.lastRenderTime;
	const timeToWait = Math.max(0, RENDER_THROTTLE_MS - timeSinceLastRender);

	// Almacenar la solicitud actual (renderFn) como la más reciente
	throttleState.pendingRequest = renderFn;

	return new Promise<T>(resolve => {
		// Agregar este resolver a la lista de resolvers pendientes
		throttleState!.pendingResolvers.push(resolve);

		// Si ya hay un timeout pendiente, NO hacer nada más
		// La solicitud actual ya fue almacenada como pendingRequest y se ejecutará con las demás
		if (throttleState!.pendingTimeout) {
			// Request encolada, esperando render pendiente
		} else {
			// No hay timeout pendiente, crear uno
			throttleState!.pendingTimeout = setTimeout(async () => {
				throttleState!.lastRenderTime = Date.now();
				throttleState!.pendingTimeout = null;

				// Ejecutar la ÚLTIMA solicitud que llegó durante el throttle window
				const lastRequest = throttleState!.pendingRequest;
				const allResolvers = throttleState!.pendingResolvers;
				throttleState!.pendingRequest = null;
				throttleState!.pendingResolvers = [];

				if (lastRequest) {
					try {
						const result = await lastRequest();
						// Resolver TODAS las promesas con el mismo resultado
						allResolvers.forEach(resolver => resolver(result));
					} catch (e) {
						console.error('[throttledRender] Error during render:', e);
						throw e;
					}
				}
			}, timeToWait);
		}
	});
}

// Función para limpieza completa de memoria antes de refresh
async function performCompleteCleanup(bodyContainer: HTMLElement, styleContainer: HTMLElement = null): Promise<void> {
	// PASO 1: Forzar recolección de basura antes de limpiar (si está disponible)
	if (typeof window !== 'undefined' && (window as any).gc) {
		(window as any).gc();
	}

	// PASO 2: Limpiar contenedores HTML (esto elimina elementos DOM)
	// Usar innerHTML = '' en lugar de cloneNode para evitar romper referencias
	bodyContainer.innerHTML = '';
	if (styleContainer) {
		styleContainer.innerHTML = '';
	}

	// PASO 3: Llamar dispose en documento anterior si existe
	const previousDoc = containerDocumentMap.get(bodyContainer);
	if (previousDoc) {
		// Obtener el renderer del documento anterior ANTES de dispose
		const rendererToDispose = typeof previousDoc.getRenderer === 'function' ? previousDoc.getRenderer() : null;
		
		// Desechar el renderer directamente si existe
		if (rendererToDispose && typeof rendererToDispose.dispose === 'function') {
			try {
				rendererToDispose.dispose();
			} catch (e) {
				// ignore
			}
		}
		
		// Luego desechar el documento
		if (typeof previousDoc.dispose === 'function') {
			previousDoc.dispose();
		}
	}

	// PASO 4: Limpiar todas las referencias de WeakMaps para este contenedor
	containerDocumentMap.delete(bodyContainer);
	containerRenderingLock.delete(bodyContainer);

	// PASO 5: Limpiar cualquier atributo data o referencias que puedan quedar en el contenedor
	if (bodyContainer.hasAttribute && bodyContainer.removeAttribute) {
		// Remover cualquier atributo data-docx que pueda haber sido agregado
		const dataAttrs = bodyContainer.attributes;
		for (let i = dataAttrs.length - 1; i >= 0; i--) {
			const attr = dataAttrs[i];
			if (attr.name.startsWith('data-docx') || attr.name.startsWith('data-')) {
				bodyContainer.removeAttribute(attr.name);
			}
		}
	}

	// PASO 6: Forzar recolección de basura
	const hasExplicitGC = typeof window !== 'undefined' && (window as any).gc;
	if (hasExplicitGC) {
		(window as any).gc();
		// Si gc() está disponible (--expose-gc en Node/Playwright), el GC fue forzado inmediatamente
		await new Promise(resolve => setTimeout(resolve, 0));
	} else {
		// Si NO tenemos gc() disponible, el navegador usa GC automático/herónico
		// Esperar más tiempo para darle oportunidad al GC de ejecutarse (reduce memory leaks)
		await new Promise(resolve => setTimeout(resolve, 50));
	}
}

// Render Synchronously
export async function renderSync(data: Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, userOptions: Partial<Options> = null): Promise<any> {
	// Aplicar throttling para garantizar espera entre renders
	return throttledRender(bodyContainer, async () => {
		// IMPORTANTE: Crear promesa del lock ANTES de hacer cualquier cosa
		let lockPromise: Promise<any> | undefined;
		const currentRendering = containerRenderingLock.get(bodyContainer);
		
		if (currentRendering) {
			lockPromise = currentRendering;
		}

		// Crear la promesa de este renderizado
		const renderingPromise = (async () => {
			// Si hay un renderizado anterior, esperar a que termine
			if (lockPromise) {
				await lockPromise;
			}

			try {
				// PASO 1: Limpieza completa antes de renderizar
				await performCompleteCleanup(bodyContainer, styleContainer);

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

		// Guardar la promesa en el lock ANTES de retornar
		containerRenderingLock.set(bodyContainer, renderingPromise);

		return renderingPromise;
	});
}

// Render Asynchronously
export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
	// Aplicar throttling para garantizar espera entre renders
	return throttledRender(bodyContainer, async () => {
		// IMPORTANTE: Crear promesa del lock ANTES de hacer cualquier cosa
		let lockPromise: Promise<any> | undefined;
		const currentRendering = containerRenderingLock.get(bodyContainer);
		
		if (currentRendering) {
			lockPromise = currentRendering;
		}

		// Crear la promesa de este renderizado
		const renderingPromise = (async () => {
			// Si hay un renderizado anterior, esperar a que termine
			if (lockPromise) {
				await lockPromise;
			}

			try {
				// PASO 1: Limpieza completa antes de renderizar
				await performCompleteCleanup(bodyContainer, styleContainer);

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

		// Guardar la promesa en el lock ANTES de retornar
		containerRenderingLock.set(bodyContainer, renderingPromise);

		return renderingPromise;
	});
}

/**
 * Limpia y libera todos los recursos asociados al contenedor proporcionado.
 * Esto incluye revocar objectURLs, destruir renderers/stages, listeners y limpiar mapas internos.
 * Útil para llamar explícitamente antes de vaciar el contenedor con `innerHTML = ''`.
 */
export async function cleanup(bodyContainer: HTMLElement, styleContainer: HTMLElement = null): Promise<void> {
	return performCompleteCleanup(bodyContainer, styleContainer);
}

// Reemplaza el documento ya parseado (o parsea si se le pasa un Blob) en el contenedor
// intentando reutilizar el renderer existente para evitar re-inicializaciones pesadas.
// - newDocOrBlob: puede ser un WordDocument ya parseado o un Blob/raw como en renderAsync
// - sync: si true usa el renderer síncrono (HtmlRendererSync), en caso contrario HtmlRenderer
/**
 * Reemplaza dinámicamente un documento en el contenedor, reutilizando opcionalmente el renderer existente.
 * Esto es más eficiente que hacer un render completo desde cero cuando se necesita cambiar el contenido.
 * 
 * @param newDocOrBlob - Un documento ya parseado (WordDocument) o un Blob/File del .docx
 * @param bodyContainer - Elemento HTML donde se renderizará
 * @param styleContainer - Elemento donde se inyectarán los estilos (por defecto null)
 * @param sync - Si true usa renderizador síncrono, si false usa asíncrono
 * @param userOptions - Opciones de configuración personalizadas
 * @returns Promise que se resuelve con el nuevo documento parseado
 * @export
 */
export async function replaceParsedDocument(newDocOrBlob: WordDocument | Blob | any, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, sync: boolean = true, userOptions?: Partial<Options>): Promise<any> {
	// merge options
	const ops = { ...defaultOptions, ...userOptions };

	// Obtener el documento previo (si existe)
	const previousDoc = containerDocumentMap.get(bodyContainer);

	// Intentar obtener el renderer del documento previo ANTES de limpiar / dispose,
	// y desasociarlo del documento para evitar que dispose() lo destruya.
	let existingRenderer: any = null;
	try {
		if (previousDoc && typeof (previousDoc as any).getRenderer === 'function') {
			existingRenderer = (previousDoc as any).getRenderer();
			// si vamos a reutilizar el renderer, desasociarlo del documento previo
			if (existingRenderer && typeof previousDoc.setRenderer === 'function') {
				try { previousDoc.setRenderer(null); } catch (e) { /* ignore */ }
			}
		}
	} catch (e) {
		// ignore
	}

	// PASO 1: limpieza completa del contenedor antes de reemplazar para evitar
	// referencias residuales (esto revoca objectURLs del documento anterior
	// y limpia el DOM/styleContainer). Similar a renderSync/renderAsync.
	await performCompleteCleanup(bodyContainer, styleContainer);

	// esperar un corto periodo para darle oportunidad al GC a liberar memoria
	// antes de crear nuevos objetos pesados (esto puede ayudar a evitar picos).
	await new Promise(resolve => setTimeout(resolve, 100));

	// Si nos pasan un Blob/raw, parsearlo a WordDocument
	let newDoc: any = null;
	if (newDocOrBlob && typeof (newDocOrBlob as any).save === 'function') {
		// ya es un WordDocument
		newDoc = newDocOrBlob;
	} else {
		newDoc = await parseAsync(newDocOrBlob, userOptions);
	}

	// Si hay un documento previo, descartarlo liberando recursos, pero no dejar que
	// elimine el renderer si vamos a reutilizarlo.
	if (previousDoc) {
		try {
			// Desasociar renderer del documento previo para evitar que dispose() lo destruya
			if (typeof previousDoc.setRenderer === 'function') {
				previousDoc.setRenderer(null);
			}
		} catch (e) {
			// ignore
		}

		try {
			if (typeof previousDoc.dispose === 'function') {
				previousDoc.dispose();
			}
		} catch (e) {
			// ignore
		}

		// limpiar el mapa para este contenedor (se volverá a establecer más abajo)
		containerDocumentMap.delete(bodyContainer);
	}

	// Determinar el renderer a usar: solo reutilizar si user explicitó reuseRenderer
	const allowReuse = ops?.reuseRenderer === true;
	let renderer: any = null;

	if (allowReuse && existingRenderer) {
		renderer = existingRenderer;
		console.warn('replaceParsedDocument: reutilizando renderer (experimental)');
	} else {
		// Si había un renderer existente, asegurarnos de liberarlo completamente
		if (existingRenderer && typeof existingRenderer.dispose === 'function') {
			try { existingRenderer.dispose(); } catch (e) { /* ignore */ }
		}

		renderer = sync ? new HtmlRendererSync() : new HtmlRenderer();
	}

	// Asociar renderer al nuevo documento para que su dispose pueda liberarlo si procede
	if (newDoc && typeof newDoc.setRenderer === 'function') {
		newDoc.setRenderer(renderer);
	}

	// Renderizar el documento (el renderer se encargará de limpiar el DOM interno)
	await renderer.render(newDoc, bodyContainer, styleContainer, ops);

	// Guardar mapping para futuras operaciones (dispose/reemplazo)
	containerDocumentMap.set(bodyContainer, newDoc);

	return newDoc;
}

// Exponer snapshot de debug
export function getDebugStats() {
	return debugStats.snapshot();
}
