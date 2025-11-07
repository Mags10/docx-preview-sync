import JSZip from 'jszip';
import { parseXmlString, XmlParser } from "../parser/xml-parser";
import { splitPath } from "../utils";
import { parseRelationships, Relationship } from "./relationship";

export interface OpenXmlPackageOptions {
    trimXmlDeclaration: boolean,
    keepOrigin: boolean,
}
//
export class OpenXmlPackage {
    xmlParser: XmlParser = new XmlParser();
    private _files: Record<string, Uint8Array> = {};
    private _zip: JSZip | null = null;

    constructor(public options: OpenXmlPackageOptions) {
    }

    private async loadAllFiles(zip: JSZip) {
        const files = zip.files;
        const promises = Object.keys(files).map(async (path) => {
            const file = files[path];
            if (!file.dir) {
                this._files[normalizePath(path)] = await file.async('uint8array');
            }
        });
        await Promise.all(promises);
        // Keep zip for save/update
    }

    get(path: string): any {
        return this._files[normalizePath(path)] ? { async: (type: JSZip.OutputType) => Promise.resolve(this.convertData(this._files[normalizePath(path)], type)) } : null;
    }

    private convertData(data: Uint8Array, type: JSZip.OutputType): any {
        switch (type) {
            case 'uint8array':
                return data;
            case 'string':
                return new TextDecoder().decode(data);
            case 'blob':
                return new Blob([data as any]);
            default:
                return data;
        }
    }

    update(path: string, content: any) {
        if (this._zip) {
            this._zip.file(path, content);
        }
    }

    static async load(input: Blob | any, options: OpenXmlPackageOptions): Promise<OpenXmlPackage> {
        let arrayBuffer: ArrayBuffer;
        if (input instanceof Blob) {
            arrayBuffer = await input.arrayBuffer();
        } else {
            arrayBuffer = input;
        }
        const zip = await JSZip.loadAsync(arrayBuffer);
        const pkg = new OpenXmlPackage(options);
        pkg._zip = zip;
        await pkg.loadAllFiles(zip);
        return pkg;
    }

    save(type: any = "blob"): Promise<any>  {
        return this._zip ? this._zip.generateAsync({ type }) : Promise.reject(new Error("Zip not loaded"));
    }

    load(path: string, type: JSZip.OutputType = "string"): Promise<any> {
        return this.get(path)?.async(type) ?? Promise.resolve(null);
    }

    async loadRelationships(path: string = null): Promise<Relationship[]> {
        let relsPath = `_rels/.rels`;

        if (path != null) {
            const [f, fn] = splitPath(path);
            relsPath = `${f}_rels/${fn}.rels`;
        }

        const txt = await this.load(relsPath);
		return txt ? parseRelationships(this.parseXmlDocument(txt).firstElementChild, this.xmlParser) : null;
    }

    /** @internal */
    parseXmlDocument(txt: string): Document {
        return parseXmlString(txt, this.options.trimXmlDeclaration);
    }
}

function normalizePath(path: string) {
    return path.startsWith('/') ? path.substr(1) : path;
}
