// 图片输入压缩：本地文件 → dataUrl → 解码 → 长边缩放 → JPEG，得到可直接发给 vision 接口的字符串
export interface CompressedImage {
  base64: string
  width: number
  height: number
}

/** 发给接口的图片长边上限：再大对理解无收益，只会拖慢上传与计费 */
const MAX_LONG_SIDE = 1536
const JPEG_QUALITY = 0.85

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result))
    reader.onerror = () => reject(new Error('图片读取失败'))
    reader.readAsDataURL(file)
  })
}

function decodeImageDataUrl(dataUrl: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve(img)
    img.onerror = () => reject(new Error('图片解码失败'))
    img.src = dataUrl
  })
}

/** 等比缩放到长边 ≤ maxSize 后转 JPEG base64 */
export function compressImage(img: HTMLImageElement, maxSize = MAX_LONG_SIDE, quality = JPEG_QUALITY): CompressedImage {
  const longSide = Math.max(img.naturalWidth, img.naturalHeight)
  const scale = longSide > maxSize ? maxSize / longSide : 1
  const width = Math.max(1, Math.round(img.naturalWidth * scale))
  const height = Math.max(1, Math.round(img.naturalHeight * scale))

  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('当前浏览器不支持 canvas')
  ctx.drawImage(img, 0, 0, width, height)

  return { base64: canvas.toDataURL('image/jpeg', quality), width, height }
}

/** 粘贴/选择的图片文件压成 dataUrl；非图片文件或解码失败直接抛错，由调用方逐条提示 */
export async function compressImageFile(file: File): Promise<string> {
  if (!file.type.startsWith('image/')) throw new Error('不是图片文件')
  const raw = await readFileAsDataUrl(file)
  const img = await decodeImageDataUrl(raw)
  return compressImage(img).base64
}
