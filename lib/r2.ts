import {
  S3Client,
  PutObjectCommand,
  DeleteObjectCommand,
  GetObjectCommand,
} from '@aws-sdk/client-s3'
import { getSignedUrl } from '@aws-sdk/s3-request-presigner'

const r2Client = new S3Client({
  region: 'auto',
  endpoint: `https://${process.env.R2_ACCOUNT_ID}.r2.cloudflarestorage.com`,
  credentials: {
    accessKeyId: process.env.R2_ACCESS_KEY_ID!,
    secretAccessKey: process.env.R2_SECRET_ACCESS_KEY!,
  },
})

const BUCKET = process.env.R2_BUCKET_NAME!

/**
 * Upload file ke R2
 */
export async function r2Upload(
  path: string,
  buffer: Buffer | Uint8Array,
  contentType: string
): Promise<void> {
  await r2Client.send(
    new PutObjectCommand({
      Bucket: BUCKET,
      Key: path,
      Body: buffer,
      ContentType: contentType,
    })
  )
}

/**
 * Download file dari R2, return Buffer
 */
export async function r2Download(path: string): Promise<Buffer> {
  const response = await r2Client.send(
    new GetObjectCommand({
      Bucket: BUCKET,
      Key: path,
    })
  )

  const chunks: Uint8Array[] = []
  for await (const chunk of response.Body as AsyncIterable<Uint8Array>) {
    chunks.push(chunk)
  }
  return Buffer.concat(chunks)
}

/**
 * Hapus file dari R2 (tidak throw kalau file tidak ada)
 */
export async function r2Delete(path: string): Promise<void> {
  try {
    await r2Client.send(
      new DeleteObjectCommand({
        Bucket: BUCKET,
        Key: path,
      })
    )
  } catch {
    // Abaikan error kalau file memang tidak ada
  }
}

/**
 * Generate presigned URL untuk akses file (default expire 1 jam)
 */
export async function r2SignedUrl(
  path: string,
  expiresInSeconds = 60 * 60
): Promise<string> {
  const command = new GetObjectCommand({
    Bucket: BUCKET,
    Key: path,
  })
  return getSignedUrl(r2Client, command, { expiresIn: expiresInSeconds })
}