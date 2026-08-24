/**
 * PptxGenJS: Media Methods
 */

import { IMG_BROKEN, SLIDE_OBJECT_TYPES } from './core-enums'
import { PresSlide, SlideLayout, ISlideRelMedia, MediaOnError } from './core-interfaces'

export type EncodeMediaOptions = {
	/** Default `'throw'` matches the historical export policy. */
	onError?: MediaOnError
}
import { applySvgFillColor, applySvgFillToDataUrl, base64ToBytes, binaryStringToBase64, bytesToBase64, importNodeBuiltin, isNodeRuntime, utf8ToBase64 } from './gen-utils'

/** Images are measured in pixels; PowerPoint slide dimensions are inches at 96 DPI */
const IMAGE_DPI = 96

type NodeMediaModules = {
	fs?: typeof import('node:fs')
	https?: typeof import('node:https')
	http?: typeof import('node:http')
}

type LoadNodeMediaModules = () => Promise<void>

/**
 * Read the natural pixel dimensions of a raster image from its header bytes.
 * Self-contained replacement for the `image-size` package (PNG/JPEG/GIF/BMP).
 * Format refs: PNG spec §11.2.2 (IHDR is always the first chunk, width/height big-endian at bytes 16-24);
 * JPEG (JFIF) SOF0-SOF15 markers carry height then width big-endian; GIF87a/89a logical-screen descriptor
 * little-endian at bytes 6-10; BMP DIB header little-endian at bytes 18-26. SVG is sized separately (gen-objects).
 *
 * Declared dims are header-only (we never decode pixels), so clamp to a sane ceiling: a maliciously crafted
 * header can claim 0xFFFFFFFF px and emit a ~44-million-inch slide (a corrupt pptx). PowerPoint's max slide is
 * 56 in (5376 px @ 96 DPI); 100 000 px is generous headroom for legit high-res sources while rejecting bombs.
 */
const MAX_IMAGE_PX = 100000

function imageDimensions(bytes: Uint8Array): { width?: number; height?: number } {
	if (bytes.length < 10) return {} // every supported header's dims start at/after byte 6
	const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	const sane = (w: number, h: number) => (w > 0 && h > 0 && w <= MAX_IMAGE_PX && h <= MAX_IMAGE_PX ? { width: w, height: h } : {})

	// PNG: 8-byte signature 89 50 4E 47 0D 0A 1A 0A, then IHDR length(4)+type(4)+width(4)+height(4)
	if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
		if (bytes.length < 24) return {}
		return sane(dv.getUint32(16, false), dv.getUint32(20, false))
	}

	// GIF: 'GIF87a'/'GIF89a', logical screen width/height little-endian at bytes 6-10
	if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
		return sane(dv.getUint16(6, true), dv.getUint16(8, true))
	}

	// BMP: 'BM', DIB width/height little-endian at bytes 18-26 (signed; height may be negative = top-down)
	if (bytes[0] === 0x42 && bytes[1] === 0x4d) {
		if (bytes.length < 26) return {}
		return sane(Math.abs(dv.getInt32(18, true)), Math.abs(dv.getInt32(22, true)))
	}

	// JPEG: FF D8, then walk marker segments until a Start-Of-Frame (C0-CF except C4/C8/CC) carries the dims
	if (bytes[0] === 0xff && bytes[1] === 0xd8) {
		let off = 2
		while (off + 9 <= bytes.length) {
			if (bytes[off] !== 0xff) { off++; continue }
			const marker = bytes[off + 1]
			const isSOF = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc
			if (isSOF) return sane(dv.getUint16(off + 7, false), dv.getUint16(off + 5, false))
			if (marker === 0xd8 || marker === 0xd9 || (marker >= 0xd0 && marker <= 0xd7)) { off += 2; continue } // no length
			const segLen = dv.getUint16(off + 2, false)
			off += 2 + segLen
		}
	}

	return {}
}

/**
 * Natural pixel size from encoded image bytes (PNG/JPEG/GIF/BMP headers).
 * Returns null when the payload is missing, not a supported raster, or claims 0×0
 * (`imageDimensions` already rejects zero and bomb-sized headers).
 *
 * contain/cover `a:srcRect` (ECMA-376 §5.1.10.55) needs the *blip* aspect, not the
 * placement box — using w/h as imgSize makes srcRect a no-op whenever they match the sizing box.
 */
export function getImagePixelSize (base64Data: unknown): { w: number, h: number } | null {
	if (typeof base64Data !== 'string' || !base64Data || base64Data === IMG_BROKEN) return null
	try {
		const dims = imageDimensions(base64ToBytes(base64Data))
		if (dims.width && dims.height) return { w: dims.width, h: dims.height }
	} catch {
		// unrecognized / truncated payload
	}
	return null
}

/**
 * Size images added without `w`/`h` to their natural dimensions (issue #34)
 * @note must run after `encodeSlideMediaRels` has resolved - it reads the encoded image bytes
 * @param {PresSlide | SlideLayout} layout - slide layout
 */
export function applyNaturalImageSizes(layout: PresSlide | SlideLayout): void {
	layout._slideObjects
		.filter(obj => obj._type === SLIDE_OBJECT_TYPES.image && obj.options?._sizeFromImage)
		.forEach(obj => {
			const strData = layout._relsMedia.find(rel => rel.rId === obj.imageRid)?.data
			if (!obj.options || typeof strData !== 'string' || !strData || strData === IMG_BROKEN) return
			// An image bound to a placeholder inherits the placeholder's w/h (slide XML resolves it); sizing it
			// to its natural pixel dims instead would collapse it to ~px/96 inches (issue #996).
			if (obj.options.placeholder) return
			try {
				const dims = imageDimensions(base64ToBytes(strData))
				if (dims.width && dims.height) {
					obj.options.w = dims.width / IMAGE_DPI
					obj.options.h = dims.height / IMAGE_DPI
				}
			} catch (_ex) {
				// Unreadable/unsupported image: keep the 1x1 inch default
			}
		})
}

/** Select relations that still need their local or remote media encoded. */
function getMediaCandidates(layout: PresSlide | SlideLayout): ISlideRelMedia[] {
	return layout._relsMedia.filter(
		rel => !rel.isLinked && rel.type !== 'online' && !rel.data && (!rel.path || (rel.path && !rel.path.includes('preencoded')))
	)
}

/**
 * Mark repeated source paths so only the first relation performs I/O.
 *
 * Mutates `isDuplicate`; subsequent helpers copy the first result to every duplicate.
 */
function markDuplicateMedia(candidates: ISlideRelMedia[]): void {
	const paths = new Set<string>()
	candidates.forEach(rel => {
		const path = rel.path ?? ''
		rel.isDuplicate = paths.has(path)
		paths.add(path)
	})
}

/** Copy an encoded or broken source relation's data to every duplicate path. */
function copyMediaToDuplicates(candidates: ISlideRelMedia[], rel: ISlideRelMedia): void {
	candidates
		.filter(duplicate => duplicate.isDuplicate && duplicate.path === rel.path)
		.forEach(duplicate => (duplicate.data = rel.data))
}

/** Mark a failed source and all duplicates with the sentinel expected by package generation. */
function markMediaBroken(candidates: ISlideRelMedia[], rel: ISlideRelMedia): void {
	rel.data = IMG_BROKEN
	copyMediaToDuplicates(candidates, rel)
}

function mediaLoadFailed(candidates: ISlideRelMedia[], rel: ISlideRelMedia, message: string, onError: MediaOnError): never | string {
	markMediaBroken(candidates, rel)
	if (onError === 'placeholder') {
		console.warn(`WARN: Unable to load image, using placeholder: ${rel.path}${message ? ` (${message})` : ''}`)
		return 'done'
	}
	throw new Error(message)
}

/**
 * Read one Node-local media source and encode it as base64.
 *
 * A read failure is recorded before rejection so duplicate references remain consistent.
 */
async function readNodeMediaFile(rel: ISlideRelMedia, candidates: ISlideRelMedia[], fs: typeof import('node:fs'), onError: MediaOnError): Promise<string> {
	try {
		const relPath = rel.path ?? ''
		const svgFill = typeof rel.fill?.color === 'string' ? rel.fill.color : undefined
		if (svgFill && (rel.type === 'image/svg+xml' || rel.extn === 'svg')) {
			const svgText = applySvgFillColor(fs.readFileSync(relPath, 'utf8'), svgFill)
			rel.data = utf8ToBase64(svgText)
		} else {
			const bitmap = fs.readFileSync(relPath)
			rel.data = bytesToBase64(new Uint8Array(bitmap))
			if (svgFill && typeof rel.data === 'string') {
				rel.data = applySvgFillToDataUrl(rel.data, svgFill)
			}
		}
		copyMediaToDuplicates(candidates, rel)
		return 'done'
	} catch (ex) {
		return mediaLoadFailed(candidates, rel, `ERROR: Unable to read media: "${rel.path}"\n${String(ex)}`, onError)
	}
}

/**
 * Load one HTTP(S) source in Node, draining error responses and handling both request and response errors.
 */
function loadNodeMediaUrl(
	rel: ISlideRelMedia,
	candidates: ISlideRelMedia[],
	https: typeof import('node:https'),
	http: typeof import('node:http'),
	onError: MediaOnError
): Promise<string> {
	const httpMod = rel.path?.startsWith('http://') ? http : https
	return new Promise<string>((resolve, reject) => {
		const fail = (message: string): void => {
			try {
				resolve(mediaLoadFailed(candidates, rel, message, onError))
			} catch (ex) {
				reject(ex)
			}
		}
		try {
			const req = httpMod.get(rel.path ?? '', res => {
				const status = res.statusCode ?? 0
				if (status < 200 || status > 299) {
					res.resume() // drain so the socket is freed
					fail(`ERROR! Unable to load image (HTTP ${status}): ${rel.path}`)
					return
				}
				let raw = ''
				res.setEncoding('binary') // IMPORTANT: Only binary encoding works
				res.on('data', chunk => (raw += chunk))
				res.on('end', () => {
					rel.data = binaryStringToBase64(raw)
					copyMediaToDuplicates(candidates, rel)
					resolve('done')
				})
				res.on('error', () => fail(`ERROR! Unable to load image (response): ${rel.path}`))
			})
			// NOTE: without this, a DNS/TLS/connection failure emits an unhandled 'error' and kills the process
			req.on('error', ex => fail(`ERROR! Unable to load image (request): ${rel.path}\n${String(ex)}`))
		} catch (ex) {
			fail(`ERROR! Unable to load image (https.get): ${rel.path}\n${String(ex)}`)
		}
	})
}

/**
 * Fetch one browser media source and convert its Blob response to a data URL.
 * SVG previews are produced only after the source relation has been populated.
 */
function loadBrowserMedia(rel: ISlideRelMedia, candidates: ISlideRelMedia[], onError: MediaOnError): Promise<string> {
	return new Promise<string>((resolve, reject) => {
		const fail = (message: string): void => {
			try {
				resolve(mediaLoadFailed(candidates, rel, message, onError))
			} catch (ex) {
				reject(ex)
			}
		}
		const xhr = new XMLHttpRequest()
		xhr.onload = () => {
			// status 0 = non-HTTP schemes (file://); anything outside 2xx is an error page, not image bytes
			if (xhr.status !== 0 && (xhr.status < 200 || xhr.status > 299)) {
				fail(`ERROR! HTTP status ${xhr.status} loading image: ${rel.path}`)
				return
			}
			const reader = new FileReader()
			reader.onloadend = () => {
				if (typeof reader.result === 'string') rel.data = reader.result
				copyMediaToDuplicates(candidates, rel)
				if (!rel.isSvgPng) resolve('done')
				else createSvgPngPreview(rel).then(() => resolve('done')).catch(reject)
			}
			reader.readAsDataURL(xhr.response)
		}
		xhr.onerror = () => {
			fail(`ERROR! Unable to load image (xhr.onerror): ${rel.path}`)
		}
		xhr.open('GET', rel.path ?? '')
		xhr.responseType = 'blob'
		xhr.send()
	})
}

/** Route one source relation to the loader valid for the active runtime. */
async function encodeMediaRelation(
	rel: ISlideRelMedia,
	candidates: ISlideRelMedia[],
	isNode: boolean,
	modules: NodeMediaModules,
	loadNodeModules: LoadNodeMediaModules,
	onError: MediaOnError
): Promise<string> {
	if (!modules.https) await loadNodeModules()

	const path = rel.path ?? ''
	if (isNode && modules.fs && !path.startsWith('http')) return readNodeMediaFile(rel, candidates, modules.fs, onError)
	if (isNode && modules.https && modules.http && path.startsWith('http')) return loadNodeMediaUrl(rel, candidates, modules.https, modules.http, onError)
	return loadBrowserMedia(rel, candidates, onError)
}

/**
 * Append SVG preview work synchronously because the caller immediately snapshots `promises`.
 * Node cannot rasterize SVG here, so it keeps the historical broken-preview sentinel behavior.
 */
function addSvgPreviewPromises(layout: PresSlide | SlideLayout, isNode: boolean, promises: Array<Promise<string>>): void {
	layout._relsMedia
		.filter(rel => rel.isSvgPng && rel.data)
		.forEach(rel => {
			// NOTE: must push synchronously - the caller snapshots this array as soon as we return
			if (isNode) {
				// SVG is not supported in Node (more info: https://github.com/gitbrent/PptxGenJS/issues/401)
				rel.data = IMG_BROKEN
				promises.push(Promise.resolve('done'))
			} else {
				promises.push(createSvgPngPreview(rel))
			}
		})
}

/**
 * Encode Image/Audio/Video into base64
 * @param {PresSlide | SlideLayout} layout - slide layout
 * @return {Promise} promise
 */
export function encodeSlideMediaRels(layout: PresSlide | SlideLayout, opts?: EncodeMediaOptions): Array<Promise<string>> {
	// STEP 1: Detect a real Node process. Browser bundles often polyfill `process`
	// (and even `process.versions.node`), so also require a document-less runtime.
	const isNode = isNodeRuntime()
	const modules: NodeMediaModules = {}

	// STEP 2: Lazy-load Node built-ins if needed. The specifier is built at runtime
	// so browser bundlers cannot statically resolve `node:https` / `https`.
	const loadNodeDeps = isNode
		? async () => {
			; ({ default: modules.fs } = await importNodeBuiltin<typeof import('node:fs')>('fs'))
			; ({ default: modules.https } = await importNodeBuiltin<typeof import('node:https')>('https'))
			; ({ default: modules.http } = await importNodeBuiltin<typeof import('node:http')>('http'))
		}
		: async () => { }
	// Immediately start it when we know we’re in Node
	if (isNode) loadNodeDeps()

	// STEP 3: Prepare promises list
	const imageProms: Array<Promise<string>> = []

	// A: Capture all audio/image/video candidates for encoding (filtering online/pre-encoded)
	const candidateRels = getMediaCandidates(layout)

	// B: PERF: Mark dupes (same `path`) to avoid loading the same media over-and-over!
	markDuplicateMedia(candidateRels)

	// STEP 4: Read/Encode each unique media item
	const onError: MediaOnError = opts?.onError === 'placeholder' ? 'placeholder' : 'throw'

	candidateRels
		.filter(rel => !rel.isDuplicate)
		.forEach(rel => imageProms.push(encodeMediaRelation(rel, candidateRels, isNode, modules, loadNodeDeps, onError)))

	// STEP 5: SVG-PNG previews
	// ......: "SVG:" base64 data still requires a png to be generated
	// ......: (`isSvgPng` flag this as the preview image, not the SVG itself)
	addSvgPreviewPromises(layout, isNode, imageProms)

	return imageProms
}

/**
 * Create SVG preview image
 * @param {ISlideRelMedia} rel - slide rel
 * @return {Promise} promise
 */
async function createSvgPngPreview(rel: ISlideRelMedia): Promise<string> {
	return await new Promise((resolve, reject) => {
		// A: Create
		const image = new Image()

		// Shared error handler (also wired to `image.onerror`); the reason string is informational only
		const handleError = (): void => {
			rel.data = IMG_BROKEN
			reject(new Error(`ERROR! Unable to load image (image.onerror): ${rel.path}`))
		}

		// B: Set onload event
		image.onload = () => {
			// First: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (image.width + image.height === 0) {
				handleError()
			}
			const canvas = document.createElement('canvas')
			const ctx = canvas.getContext('2d')
			if (!ctx) {
				handleError()
				return
			}
			canvas.width = image.width
			canvas.height = image.height
			ctx.drawImage(image, 0, 0)
			// Users running on local machine will get the following error:
			// "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
			// when the canvas.toDataURL call executes below.
			try {
				rel.data = canvas.toDataURL(rel.type)
				resolve('done')
			} catch (_ex) {
				handleError()
			}
		}
		image.onerror = handleError

		// C: Load image
		image.src = typeof rel.data === 'string' ? rel.data : IMG_BROKEN
	})
}
