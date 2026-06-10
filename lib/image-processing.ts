import sharp from 'sharp'

export type ProcessResult =
  | { ok: true; buffer: Buffer; width: number; height: number }
  | { ok: false; reason: string }

/**
 * Detect blur using variance of Laplacian method.
 * Lower variance = more blurry.
 */
export async function detectBlur(buffer: Buffer): Promise<number> {
  // Convert to grayscale, get raw pixels
  const { data, info } = await sharp(buffer)
    .greyscale()
    .resize(400, 400, { fit: 'inside' })
    .raw()
    .toBuffer({ resolveWithObject: true })

  const { width, height } = info
  const pixels = new Uint8Array(data)

  // Apply Laplacian kernel: [0,1,0 / 1,-4,1 / 0,1,0]
  let variance = 0
  let mean = 0
  const values: number[] = []

  for (let y = 1; y < height - 1; y++) {
    for (let x = 1; x < width - 1; x++) {
      const idx = y * width + x
      const lap =
        pixels[idx - width] +
        pixels[idx + width] +
        pixels[idx - 1] +
        pixels[idx + 1] -
        4 * pixels[idx]
      values.push(lap)
      mean += lap
    }
  }

  mean /= values.length
  for (const v of values) variance += (v - mean) ** 2
  variance /= values.length

  return variance
}

/**
 * Process a receipt image:
 * 1. Check blur
 * 2. Auto-crop to content (trim background)
 * 3. Resize to standard size (A5 portrait equivalent)
 * 4. Convert to JPEG
 */
export async function processReceiptImage(
  inputBuffer: Buffer,
  blurThreshold = 100
): Promise<ProcessResult> {
  // Check file size - max 10MB
  if (inputBuffer.length > 10 * 1024 * 1024) {
    return { ok: false, reason: 'File terlalu besar (max 10MB)' }
  }

  // Validate it's actually an image
  let metadata
  try {
    metadata = await sharp(inputBuffer).metadata()
  } catch {
    return { ok: false, reason: 'File bukan gambar yang valid' }
  }

  if (!metadata.width || !metadata.height) {
    return { ok: false, reason: 'Gambar tidak bisa dibaca' }
  }

  // Check blur
  const blurScore = await detectBlur(inputBuffer)
  if (blurScore < blurThreshold) {
    return {
      ok: false,
      reason: `Foto terlalu blur (skor: ${blurScore.toFixed(1)}, minimum: ${blurThreshold}). Silakan foto ulang dengan lebih jelas.`
    }
  }

  // Process image:
  // - Auto-trim whitespace/background edges
  // - Resize to fit within 800x1100 (receipt-like proportion)
  // - Add small white padding
  // - Convert to JPEG quality 85
  try {
    const processed = await sharp(inputBuffer)
      .rotate() // auto-rotate based on EXIF
      .trim({
        background: { r: 255, g: 255, b: 255, alpha: 1 },
        threshold: 40
      })
      .resize(800, 1100, {
        fit: 'inside',
        withoutEnlargement: false
      })
      .extend({
        top: 20,
        bottom: 20,
        left: 20,
        right: 20,
        background: { r: 255, g: 255, b: 255, alpha: 1 }
      })
      .jpeg({ quality: 85, progressive: true })
      .toBuffer()

    const finalMeta = await sharp(processed).metadata()

    return {
      ok: true,
      buffer: processed,
      width: finalMeta.width || 800,
      height: finalMeta.height || 1100
    }
  } catch (err) {
    return { ok: false, reason: 'Gagal memproses gambar' }
  }
}

/**
 * Create a thumbnail version for grid preview
 */
export async function createThumbnail(buffer: Buffer): Promise<Buffer> {
  return sharp(buffer)
    .resize(300, 400, { fit: 'cover', position: 'centre' })
    .jpeg({ quality: 70 })
    .toBuffer()
}
