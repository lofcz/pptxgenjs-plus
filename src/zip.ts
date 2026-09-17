import type { JSZip } from '@node-projects/jszip'

/** Add an OPC part without creating an explicit ZIP directory entry for its path. */
export function addZipFile<T extends JSZip.InputType> (
	zip: JSZip,
	path: string,
	data: InputByType[T] | Promise<InputByType[T]>,
	options: JSZip.JSZipFileOptions = {},
): void {
	zip.file(path, data, { ...options, createFolders: false })
}
