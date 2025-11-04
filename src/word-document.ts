import { OutputType } from "jszip";
import mime from './mime';

import { DocumentParser } from './document-parser';
import { Relationship, RelationshipTypes } from './common/relationship';
import { Part } from './common/part';
import { FontTablePart } from './font-table/font-table';
import { OpenXmlPackage } from './common/open-xml-package';
import { DocumentPart } from './document/document-part';
import { blobToBase64, resolvePath, splitPath } from './utils';
import { NumberingPart } from './numbering/numbering-part';
import { StylesPart } from './styles/styles-part';
import { FooterPart, HeaderPart } from "./header-footer/parts";
import { ExtendedPropsPart } from "./document-props/extended-props-part";
import { CorePropsPart } from "./document-props/core-props-part";
import { ThemePart } from "./theme/theme-part";
import { EndnotesPart, FootnotesPart } from "./notes/parts";
import { SettingsPart } from "./settings/settings-part";
import { CustomPropsPart } from "./document-props/custom-props-part";
import { CommentsPart } from "./comments/comments-part";

const topLevelRels = [
	{ type: RelationshipTypes.OfficeDocument, target: "word/document.xml" },
	{ type: RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
	{ type: RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
	{ type: RelationshipTypes.CustomProperties, target: "docProps/custom.xml" },
];

// word文件解析器：blob对象 => Object对象
export class WordDocument {
	private _package: OpenXmlPackage;
	private _parser: DocumentParser;
	private _options: any;

	rels: Relationship[];
	parts: Part[] = [];
	partsMap: Record<string, Part> = {};

	documentPart: DocumentPart;
	fontTablePart: FontTablePart;
	numberingPart: NumberingPart;
	stylesPart: StylesPart;
	footnotesPart: FootnotesPart;
	endnotesPart: EndnotesPart;
	themePart: ThemePart;
	corePropsPart: CorePropsPart;
	extendedPropsPart: ExtendedPropsPart;
	settingsPart: SettingsPart;
	commentsPart: CommentsPart;

	// Lista de URLs creados con createObjectURL que necesitan ser revocados
	private createdObjectURLs: string[] = [];

	// Renderer actual para limpieza
	private _renderer: any = null;

	static async load(blob: Blob | any, parser: DocumentParser, options: any): Promise<WordDocument> {
		var d = new WordDocument();

		d._options = options;
		d._parser = parser;
		// 解压缩word文件转换为Office Open XML
		d._package = await OpenXmlPackage.load(blob, options);
		d.rels = await d._package.loadRelationships();

		await Promise.all(topLevelRels.map(rel => {
			const r = d.rels.find(x => x.type === rel.type) ?? rel; //fallback
			return d.loadRelationshipPart(r.target, r.type);
		}));

		return d;
	}

	save(type = "blob"): Promise<any> {
		return this._package.save(type);
	}

	// Establecer el renderer para limpieza posterior
	setRenderer(renderer: any) {
		this._renderer = renderer;
	}

	dispose() {
		// PASO 1: Limpiar referencia al renderer primero para romper el ciclo
		const renderer = this._renderer;
		this._renderer = null;

		// PASO 2: Liberar el renderer si existe (esto debería limpiar todas las referencias DOM)
		if (renderer && typeof renderer.dispose === 'function') {
			renderer.dispose();
		}

		// PASO 3: Liberar el parser si tiene método dispose
		if (this._parser && typeof this._parser.dispose === 'function') {
			this._parser.dispose();
		}

		// PASO 4: Revocar todos los URLs de objetos creados para liberar memoria
		for (const url of this.createdObjectURLs) {
			try {
				URL.revokeObjectURL(url);
			} catch (e) {
				// Ignorar errores si el URL ya fue revocado
			}
		}
		this.createdObjectURLs = [];

		// PASO 5: Limpiar todas las referencias de partes específicas
		this.documentPart = null;
		this.fontTablePart = null;
		this.numberingPart = null;
		this.stylesPart = null;
		this.footnotesPart = null;
		this.endnotesPart = null;
		this.themePart = null;
		this.corePropsPart = null;
		this.extendedPropsPart = null;
		this.settingsPart = null;
		this.commentsPart = null;

		// PASO 6: Limpiar arrays y mapas
		this.parts = null;
		this.partsMap = null;
		this.rels = null;

		// PASO 7: Limpiar referencias al package y parser
		this._package = null;
		this._parser = null;
		this._options = null;
	}

	private async loadRelationshipPart(path: string, type: string): Promise<Part> {
		if (this.partsMap[path])
			return this.partsMap[path];

		if (!this._package.get(path))
			return null;

		let part: Part = null;

		switch (type) {
			case RelationshipTypes.OfficeDocument:
				this.documentPart = part = new DocumentPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.FontTable:
				this.fontTablePart = part = new FontTablePart(this._package, path);
				break;

			case RelationshipTypes.Numbering:
				this.numberingPart = part = new NumberingPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Styles:
				this.stylesPart = part = new StylesPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Theme:
				this.themePart = part = new ThemePart(this._package, path);
				break;

			case RelationshipTypes.Footnotes:
				this.footnotesPart = part = new FootnotesPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Endnotes:
				this.endnotesPart = part = new EndnotesPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Footer:
				part = new FooterPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.Header:
				part = new HeaderPart(this._package, path, this._parser);
				break;

			case RelationshipTypes.CoreProperties:
				this.corePropsPart = part = new CorePropsPart(this._package, path);
				break;

			case RelationshipTypes.ExtendedProperties:
				this.extendedPropsPart = part = new ExtendedPropsPart(this._package, path);
				break;

			case RelationshipTypes.CustomProperties:
				part = new CustomPropsPart(this._package, path);
				break;

			case RelationshipTypes.Settings:
				this.settingsPart = part = new SettingsPart(this._package, path);
				break;

			case RelationshipTypes.Comments:
				this.commentsPart = part = new CommentsPart(this._package, path, this._parser);
				break;
		}

		if (part == null)
			return Promise.resolve(null);

		this.partsMap[path] = part;
		this.parts.push(part);

		await part.load();

		if (part.rels?.length > 0) {
			const [folder] = splitPath(part.path);
			await Promise.all(part.rels.map(rel => this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type)));
		}

		return part;
	}

	async loadDocumentImage(id: string, part?: Part): Promise<string> {
		const blob = await this.loadResource(part ?? this.documentPart, id, "blob");
		return this.blobToURL(blob);
	}

	async loadNumberingImage(id: string): Promise<string> {
		const blob = await this.loadResource(this.numberingPart, id, "blob");
		return this.blobToURL(blob);
	}

	async loadFont(id: string, key: string): Promise<string> {
		const x = await this.loadResource(this.fontTablePart, id, "uint8array");
		return x ? this.blobToURL(new Blob([deobfuscate(x, key) as any])) : x;
	}

	private blobToURL(blob: Blob): string | Promise<string> {
		if (!blob)
			return null;

		if (this._options.useBase64URL) {
			return blobToBase64(blob);
		}

		const url = URL.createObjectURL(blob);
		// Trackear el URL creado para poder revocarlo después
		this.createdObjectURLs.push(url);
		return url;
	}

	findPartByRelId(id: string, documentPart: Part = null) {
		var rel = (documentPart.rels ?? this.rels).find(r => r.id == id);
		const folder = documentPart ? splitPath(documentPart.path)[0] : '';
		return rel ? this.partsMap[resolvePath(rel.target, folder)] : null;
	}

	getPathById(part: Part, id: string): string {
		const rel = part.rels.find(x => x.id == id);
		const [folder] = splitPath(part.path);
		return rel ? resolvePath(rel.target, folder) : null;
	}


	private async loadResource(part: Part, id: string, outputType: OutputType) {
		const path = this.getPathById(part, id);
		// TODO 暂时使用文件扩展名推断MIME类型，实际上并不准确
		let type = mime.getType(path);
		if (path) {
			if (outputType === "blob") {
				let data = await this._package.load(path, "uint8array");
				return new Blob([data], { type });
			} else {
				return await this._package.load(path, outputType);
			}
		} else {
			return Promise.resolve(null);
		}
	}
}

export function deobfuscate(data: Uint8Array, guidKey: string): Uint8Array {
	const len = 16;
	const trimmed = guidKey.replace(/{|}|-/g, "");
	const numbers = new Array(len);

	for (let i = 0; i < len; i++)
		numbers[len - i - 1] = parseInt(trimmed.substr(i * 2, 2), 16);

	for (let i = 0; i < 32; i++)
		data[i] = data[i] ^ numbers[i % len]

	return data;
}
