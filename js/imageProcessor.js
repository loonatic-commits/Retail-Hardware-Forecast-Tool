/**
 * SnapVault - Image Processing Engine
 * Canvas-based image manipulation: adjustments, filters, transforms, and auto-enhance.
 */
class ImageProcessor {
  constructor() {
    this.filters = {
      none:      { name: 'Original',  brightness: 0, contrast: 0, saturation: 0, warmth: 0, sepia: 0, grayscale: 0 },
      vivid:     { name: 'Vivid',     brightness: 10, contrast: 20, saturation: 30, warmth: 10, sepia: 0, grayscale: 0 },
      warm:      { name: 'Warm',      brightness: 5, contrast: 5, saturation: 10, warmth: 40, sepia: 0, grayscale: 0 },
      cool:      { name: 'Cool',      brightness: 5, contrast: 10, saturation: -10, warmth: -30, sepia: 0, grayscale: 0 },
      vintage:   { name: 'Vintage',   brightness: 5, contrast: -10, saturation: -20, warmth: 20, sepia: 40, grayscale: 0 },
      bw:        { name: 'B&W',       brightness: 10, contrast: 20, saturation: 0, warmth: 0, sepia: 0, grayscale: 100 },
      sepia:     { name: 'Sepia',     brightness: 5, contrast: 0, saturation: 0, warmth: 0, sepia: 80, grayscale: 0 },
      dramatic:  { name: 'Dramatic',  brightness: -5, contrast: 40, saturation: 10, warmth: 0, sepia: 0, grayscale: 0 },
      fade:      { name: 'Fade',      brightness: 15, contrast: -15, saturation: -20, warmth: 5, sepia: 10, grayscale: 0 },
      noir:      { name: 'Noir',      brightness: -5, contrast: 30, saturation: 0, warmth: 0, sepia: 0, grayscale: 100 },
      summer:    { name: 'Summer',    brightness: 10, contrast: 10, saturation: 20, warmth: 25, sepia: 0, grayscale: 0 },
      winter:    { name: 'Winter',    brightness: 5, contrast: 15, saturation: -15, warmth: -20, sepia: 0, grayscale: 0 },
    };
  }

  /**
   * Load an image from a data URL.
   */
  loadImage(dataUrl) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = dataUrl;
    });
  }

  /**
   * Apply adjustments to pixel data on canvas.
   * @param {CanvasRenderingContext2D} ctx
   * @param {number} width
   * @param {number} height
   * @param {Object} adjustments - { brightness, contrast, saturation, warmth, sharpness, vignette, sepia, grayscale }
   */
  applyAdjustments(ctx, width, height, adj) {
    const imageData = ctx.getImageData(0, 0, width, height);
    const data = imageData.data;

    const brightness = (adj.brightness || 0) * 2.55;
    const contrast = (adj.contrast || 0) / 100;
    const saturation = (adj.saturation || 0) / 100;
    const warmth = (adj.warmth || 0) * 2.55;
    const sepia = (adj.sepia || 0) / 100;
    const grayscale = (adj.grayscale || 0) / 100;

    const contrastFactor = (259 * (contrast * 255 + 255)) / (255 * (259 - contrast * 255));

    for (let i = 0; i < data.length; i += 4) {
      let r = data[i];
      let g = data[i + 1];
      let b = data[i + 2];

      // Brightness
      r += brightness;
      g += brightness;
      b += brightness;

      // Contrast
      r = contrastFactor * (r - 128) + 128;
      g = contrastFactor * (g - 128) + 128;
      b = contrastFactor * (b - 128) + 128;

      // Warmth (shift red/blue)
      r += warmth * 0.5;
      b -= warmth * 0.3;

      // Saturation
      const gray = 0.2126 * r + 0.7152 * g + 0.0722 * b;
      r = gray + (r - gray) * (1 + saturation);
      g = gray + (g - gray) * (1 + saturation);
      b = gray + (b - gray) * (1 + saturation);

      // Grayscale
      if (grayscale > 0) {
        const grayVal = 0.2126 * r + 0.7152 * g + 0.0722 * b;
        r = r * (1 - grayscale) + grayVal * grayscale;
        g = g * (1 - grayscale) + grayVal * grayscale;
        b = b * (1 - grayscale) + grayVal * grayscale;
      }

      // Sepia
      if (sepia > 0) {
        const sr = 0.393 * r + 0.769 * g + 0.189 * b;
        const sg = 0.349 * r + 0.686 * g + 0.168 * b;
        const sb = 0.272 * r + 0.534 * g + 0.131 * b;
        r = r * (1 - sepia) + sr * sepia;
        g = g * (1 - sepia) + sg * sepia;
        b = b * (1 - sepia) + sb * sepia;
      }

      data[i] = Math.max(0, Math.min(255, r));
      data[i + 1] = Math.max(0, Math.min(255, g));
      data[i + 2] = Math.max(0, Math.min(255, b));
    }

    ctx.putImageData(imageData, 0, 0);

    // Vignette (drawn on top)
    if (adj.vignette && adj.vignette > 0) {
      this.applyVignette(ctx, width, height, adj.vignette / 100);
    }

    // Sharpness (convolution)
    if (adj.sharpness && adj.sharpness > 0) {
      this.applySharpen(ctx, width, height, adj.sharpness / 100);
    }
  }

  /**
   * Apply vignette effect.
   */
  applyVignette(ctx, width, height, amount) {
    const cx = width / 2;
    const cy = height / 2;
    const radius = Math.max(cx, cy);
    const gradient = ctx.createRadialGradient(cx, cy, radius * 0.3, cx, cy, radius);
    gradient.addColorStop(0, 'rgba(0,0,0,0)');
    gradient.addColorStop(1, `rgba(0,0,0,${amount * 0.7})`);
    ctx.fillStyle = gradient;
    ctx.fillRect(0, 0, width, height);
  }

  /**
   * Simple sharpen using unsharp mask approximation.
   */
  applySharpen(ctx, width, height, amount) {
    const imageData = ctx.getImageData(0, 0, width, height);
    const data = imageData.data;
    const copy = new Uint8ClampedArray(data);
    const w = width * 4;
    const strength = amount * 1.5;

    for (let y = 1; y < height - 1; y++) {
      for (let x = 1; x < width - 1; x++) {
        const idx = (y * width + x) * 4;
        for (let c = 0; c < 3; c++) {
          const center = copy[idx + c] * 5;
          const neighbors =
            copy[idx - w + c] +
            copy[idx + w + c] +
            copy[idx - 4 + c] +
            copy[idx + 4 + c];
          const sharpened = center - neighbors;
          data[idx + c] = Math.max(0, Math.min(255,
            copy[idx + c] + (sharpened - copy[idx + c]) * strength
          ));
        }
      }
    }
    ctx.putImageData(imageData, 0, 0);
  }

  /**
   * Auto-enhance: analyze image and apply optimal adjustments.
   */
  autoEnhance(ctx, width, height) {
    const imageData = ctx.getImageData(0, 0, width, height);
    const data = imageData.data;

    // Analyze image statistics
    let totalR = 0, totalG = 0, totalB = 0;
    let minBrightness = 255, maxBrightness = 0;
    let totalSaturation = 0;
    const pixelCount = data.length / 4;
    const sampleStep = Math.max(1, Math.floor(pixelCount / 10000)); // Sample for speed

    let samples = 0;
    for (let i = 0; i < data.length; i += 4 * sampleStep) {
      const r = data[i], g = data[i+1], b = data[i+2];
      const brightness = (r + g + b) / 3;
      totalR += r; totalG += g; totalB += b;
      minBrightness = Math.min(minBrightness, brightness);
      maxBrightness = Math.max(maxBrightness, brightness);

      const max = Math.max(r, g, b);
      const min = Math.min(r, g, b);
      totalSaturation += max > 0 ? (max - min) / max : 0;
      samples++;
    }

    const avgR = totalR / samples;
    const avgG = totalG / samples;
    const avgB = totalB / samples;
    const avgBrightness = (avgR + avgG + avgB) / 3;
    const avgSaturation = totalSaturation / samples;
    const dynamicRange = maxBrightness - minBrightness;

    // Calculate corrections
    const adjustments = {
      brightness: 0,
      contrast: 0,
      saturation: 0,
      warmth: 0,
      sharpness: 20,
      vignette: 0,
      sepia: 0,
      grayscale: 0,
    };

    // Brightness: target around 128
    if (avgBrightness < 100) adjustments.brightness = Math.min(30, (128 - avgBrightness) * 0.3);
    else if (avgBrightness > 160) adjustments.brightness = Math.max(-20, (128 - avgBrightness) * 0.2);

    // Contrast: expand dynamic range if narrow
    if (dynamicRange < 180) adjustments.contrast = Math.min(25, (200 - dynamicRange) * 0.15);

    // Saturation: boost if dull
    if (avgSaturation < 0.3) adjustments.saturation = Math.min(25, (0.4 - avgSaturation) * 60);

    // Warmth: slight warming for cold images
    const colorTemp = avgR - avgB;
    if (colorTemp < -10) adjustments.warmth = Math.min(15, (-colorTemp) * 0.3);

    return adjustments;
  }

  /**
   * Generate a thumbnail from a data URL.
   */
  async generateThumbnail(dataUrl, maxSize = 300) {
    const img = await this.loadImage(dataUrl);
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');

    let { width, height } = img;
    if (width > height) {
      if (width > maxSize) { height *= maxSize / width; width = maxSize; }
    } else {
      if (height > maxSize) { width *= maxSize / height; height = maxSize; }
    }

    canvas.width = width;
    canvas.height = height;
    ctx.drawImage(img, 0, 0, width, height);
    return canvas.toDataURL('image/jpeg', 0.7);
  }

  /**
   * Read a File object as a data URL.
   */
  readFileAsDataURL(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  /**
   * Generate a filter preview thumbnail.
   */
  async generateFilterPreview(dataUrl, filterKey, size = 80) {
    const img = await this.loadImage(dataUrl);
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    canvas.width = size;
    canvas.height = size;

    // Draw cropped square
    const min = Math.min(img.width, img.height);
    const sx = (img.width - min) / 2;
    const sy = (img.height - min) / 2;
    ctx.drawImage(img, sx, sy, min, min, 0, 0, size, size);

    // Apply filter
    const filter = this.filters[filterKey];
    if (filter && filterKey !== 'none') {
      this.applyAdjustments(ctx, size, size, filter);
    }

    return canvas;
  }

  /**
   * Rotate image by degrees (90, 180, 270).
   */
  rotateImage(canvas, ctx, img, degrees) {
    const rad = (degrees * Math.PI) / 180;
    if (degrees === 90 || degrees === 270) {
      canvas.width = img.height;
      canvas.height = img.width;
    } else {
      canvas.width = img.width;
      canvas.height = img.height;
    }
    ctx.save();
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate(rad);
    ctx.drawImage(img, -img.width / 2, -img.height / 2);
    ctx.restore();
  }

  /**
   * Flip image horizontally or vertically.
   */
  flipImage(canvas, ctx, img, horizontal = true) {
    canvas.width = img.width;
    canvas.height = img.height;
    ctx.save();
    if (horizontal) {
      ctx.translate(img.width, 0);
      ctx.scale(-1, 1);
    } else {
      ctx.translate(0, img.height);
      ctx.scale(1, -1);
    }
    ctx.drawImage(img, 0, 0);
    ctx.restore();
  }
}

const imageProcessor = new ImageProcessor();
