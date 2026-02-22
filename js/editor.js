/**
 * SnapVault - Photo Editor
 * Full-featured photo editor with adjustments, filters, crop, and transforms.
 */
class PhotoEditor {
  constructor() {
    this.canvas = document.getElementById('editor-canvas');
    this.ctx = this.canvas.getContext('2d');
    this.originalImage = null;
    this.currentPhoto = null;
    this.rotation = 0;
    this.flipH = false;
    this.flipV = false;
    this.currentFilter = 'none';

    this.adjustments = {
      brightness: 0,
      contrast: 0,
      saturation: 0,
      warmth: 0,
      sharpness: 0,
      vignette: 0,
      sepia: 0,
      grayscale: 0,
    };

    // Crop state
    this.cropMode = false;
    this.cropRect = null;
    this.cropRatio = 'free';
    this.cropOverlay = null;

    this._bindEvents();
  }

  _bindEvents() {
    // Tool tabs
    document.querySelectorAll('.tool-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.tool-tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tool-content').forEach(p => p.classList.remove('active'));
        tab.classList.add('active');
        const toolId = tab.dataset.tool + '-panel';
        document.getElementById(toolId).classList.add('active');

        // Enable/disable crop mode
        if (tab.dataset.tool === 'crop') {
          this._enableCropMode();
        } else {
          this._disableCropMode();
        }
      });
    });

    // Adjustment sliders
    const sliderIds = ['brightness', 'contrast', 'saturation', 'warmth', 'sharpness', 'vignette'];
    sliderIds.forEach(id => {
      const slider = document.getElementById(id);
      if (slider) {
        slider.addEventListener('input', () => {
          this.adjustments[id] = parseInt(slider.value);
          slider.nextElementSibling.textContent = slider.value;
          this._renderPreview();
        });
      }
    });

    // Auto enhance
    document.getElementById('auto-enhance-btn').addEventListener('click', () => {
      this._autoEnhance();
    });

    // Transform buttons
    document.getElementById('rotate-left').addEventListener('click', () => {
      this.rotation = (this.rotation - 90 + 360) % 360;
      this._renderPreview();
    });

    document.getElementById('rotate-right').addEventListener('click', () => {
      this.rotation = (this.rotation + 90) % 360;
      this._renderPreview();
    });

    document.getElementById('flip-h').addEventListener('click', () => {
      this.flipH = !this.flipH;
      this._renderPreview();
    });

    document.getElementById('flip-v').addEventListener('click', () => {
      this.flipV = !this.flipV;
      this._renderPreview();
    });

    // Crop ratios
    document.querySelectorAll('.crop-ratio').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.crop-ratio').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        this.cropRatio = btn.dataset.ratio;
        if (this.cropMode) {
          this._updateCropOverlay();
        }
      });
    });

    // Crop apply
    document.getElementById('crop-apply').addEventListener('click', () => {
      this._applyCrop();
    });

    // Header buttons
    document.getElementById('editor-cancel').addEventListener('click', () => {
      this.close();
    });

    document.getElementById('editor-reset').addEventListener('click', () => {
      this._reset();
    });

    document.getElementById('editor-save').addEventListener('click', () => {
      this._save();
    });
  }

  async open(photo) {
    this.currentPhoto = photo;
    this.rotation = 0;
    this.flipH = false;
    this.flipV = false;
    this.currentFilter = 'none';
    this.adjustments = { brightness: 0, contrast: 0, saturation: 0, warmth: 0, sharpness: 0, vignette: 0, sepia: 0, grayscale: 0 };
    this.cropMode = false;

    // Reset slider UI
    ['brightness', 'contrast', 'saturation', 'warmth', 'sharpness', 'vignette'].forEach(id => {
      const slider = document.getElementById(id);
      if (slider) {
        slider.value = 0;
        slider.nextElementSibling.textContent = '0';
      }
    });

    // Reset tabs
    document.querySelectorAll('.tool-tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tool-content').forEach(p => p.classList.remove('active'));
    document.querySelector('.tool-tab[data-tool="adjust"]').classList.add('active');
    document.getElementById('adjust-panel').classList.add('active');

    this.originalImage = await imageProcessor.loadImage(photo.dataUrl);
    this._renderPreview();
    this._generateFilterPreviews();

    document.getElementById('photo-editor').classList.remove('hidden');
  }

  close() {
    this._disableCropMode();
    document.getElementById('photo-editor').classList.add('hidden');
    this.originalImage = null;
    this.currentPhoto = null;
  }

  _renderPreview() {
    if (!this.originalImage) return;

    const img = this.originalImage;
    let w = img.width;
    let h = img.height;

    // Handle rotation dimensions
    if (this.rotation === 90 || this.rotation === 270) {
      this.canvas.width = h;
      this.canvas.height = w;
    } else {
      this.canvas.width = w;
      this.canvas.height = h;
    }

    this.ctx.save();
    this.ctx.translate(this.canvas.width / 2, this.canvas.height / 2);

    // Apply rotation
    this.ctx.rotate((this.rotation * Math.PI) / 180);

    // Apply flips
    const sx = this.flipH ? -1 : 1;
    const sy = this.flipV ? -1 : 1;
    this.ctx.scale(sx, sy);

    this.ctx.drawImage(img, -w / 2, -h / 2, w, h);
    this.ctx.restore();

    // Apply filter adjustments first
    const filterAdj = this.currentFilter !== 'none'
      ? { ...imageProcessor.filters[this.currentFilter] }
      : {};

    // Merge manual adjustments on top
    const merged = {
      brightness: (filterAdj.brightness || 0) + this.adjustments.brightness,
      contrast: (filterAdj.contrast || 0) + this.adjustments.contrast,
      saturation: (filterAdj.saturation || 0) + this.adjustments.saturation,
      warmth: (filterAdj.warmth || 0) + this.adjustments.warmth,
      sharpness: this.adjustments.sharpness,
      vignette: this.adjustments.vignette,
      sepia: filterAdj.sepia || 0,
      grayscale: filterAdj.grayscale || 0,
    };

    const hasAdjustments = Object.values(merged).some(v => v !== 0);
    if (hasAdjustments) {
      imageProcessor.applyAdjustments(this.ctx, this.canvas.width, this.canvas.height, merged);
    }
  }

  async _generateFilterPreviews() {
    const grid = document.getElementById('filters-grid');
    grid.innerHTML = '';

    for (const [key, filter] of Object.entries(imageProcessor.filters)) {
      const item = document.createElement('div');
      item.className = 'filter-item' + (key === this.currentFilter ? ' active' : '');
      item.dataset.filter = key;

      const previewCanvas = await imageProcessor.generateFilterPreview(
        this.currentPhoto.dataUrl, key, 80
      );
      previewCanvas.className = 'filter-preview';

      const name = document.createElement('span');
      name.className = 'filter-name';
      name.textContent = filter.name;

      item.appendChild(previewCanvas);
      item.appendChild(name);

      item.addEventListener('click', () => {
        document.querySelectorAll('.filter-item').forEach(fi => fi.classList.remove('active'));
        item.classList.add('active');
        this.currentFilter = key;
        this._renderPreview();
      });

      grid.appendChild(item);
    }
  }

  _autoEnhance() {
    // Render clean image first
    const tempCanvas = document.createElement('canvas');
    const tempCtx = tempCanvas.getContext('2d');
    tempCanvas.width = this.originalImage.width;
    tempCanvas.height = this.originalImage.height;
    tempCtx.drawImage(this.originalImage, 0, 0);

    const suggested = imageProcessor.autoEnhance(tempCtx, tempCanvas.width, tempCanvas.height);

    // Apply to sliders
    this.adjustments = { ...suggested };
    ['brightness', 'contrast', 'saturation', 'warmth', 'sharpness', 'vignette'].forEach(id => {
      const slider = document.getElementById(id);
      if (slider) {
        slider.value = Math.round(this.adjustments[id]);
        slider.nextElementSibling.textContent = Math.round(this.adjustments[id]);
      }
    });

    this._renderPreview();
    app.showToast('Auto enhance applied');
  }

  _reset() {
    this.rotation = 0;
    this.flipH = false;
    this.flipV = false;
    this.currentFilter = 'none';
    this.adjustments = { brightness: 0, contrast: 0, saturation: 0, warmth: 0, sharpness: 0, vignette: 0, sepia: 0, grayscale: 0 };

    ['brightness', 'contrast', 'saturation', 'warmth', 'sharpness', 'vignette'].forEach(id => {
      const slider = document.getElementById(id);
      if (slider) {
        slider.value = 0;
        slider.nextElementSibling.textContent = '0';
      }
    });

    document.querySelectorAll('.filter-item').forEach(fi => {
      fi.classList.toggle('active', fi.dataset.filter === 'none');
    });

    this._renderPreview();
    app.showToast('Reset to original');
  }

  async _save() {
    if (!this.currentPhoto) return;

    // The current canvas state is the final image
    const dataUrl = this.canvas.toDataURL('image/jpeg', 0.92);
    const thumbnail = await imageProcessor.generateThumbnail(dataUrl);

    this.currentPhoto.dataUrl = dataUrl;
    this.currentPhoto.thumbnail = thumbnail;
    this.currentPhoto.editedAt = Date.now();

    await db.updatePhoto(this.currentPhoto);
    this.close();
    app.showToast('Photo saved');
    app.refreshCurrentView();
  }

  // ---- Crop ----

  _enableCropMode() {
    this.cropMode = true;
    this._createCropOverlay();
  }

  _disableCropMode() {
    this.cropMode = false;
    if (this.cropOverlay) {
      this.cropOverlay.remove();
      this.cropOverlay = null;
    }
  }

  _createCropOverlay() {
    if (this.cropOverlay) this.cropOverlay.remove();

    const canvasArea = document.querySelector('.editor-canvas-area');
    const canvasRect = this.canvas.getBoundingClientRect();

    const overlay = document.createElement('div');
    overlay.className = 'crop-overlay';
    overlay.style.position = 'absolute';

    // Initial crop: full canvas
    const margin = 20;
    overlay.style.left = (canvasRect.left - canvasArea.getBoundingClientRect().left + margin) + 'px';
    overlay.style.top = (canvasRect.top - canvasArea.getBoundingClientRect().top + margin) + 'px';
    overlay.style.width = (canvasRect.width - margin * 2) + 'px';
    overlay.style.height = (canvasRect.height - margin * 2) + 'px';

    // Add resize handles
    ['nw', 'ne', 'sw', 'se'].forEach(pos => {
      const handle = document.createElement('div');
      handle.className = `crop-handle ${pos}`;
      overlay.appendChild(handle);
    });

    canvasArea.appendChild(overlay);
    this.cropOverlay = overlay;

    this._makeCropDraggable(overlay, canvasRect, canvasArea.getBoundingClientRect());
  }

  _makeCropDraggable(overlay, canvasRect, areaRect) {
    let isDragging = false;
    let isResizing = false;
    let resizeHandle = '';
    let startX, startY, startLeft, startTop, startWidth, startHeight;

    const onMouseDown = (e) => {
      e.preventDefault();
      const target = e.target;
      startX = e.clientX;
      startY = e.clientY;
      startLeft = overlay.offsetLeft;
      startTop = overlay.offsetTop;
      startWidth = overlay.offsetWidth;
      startHeight = overlay.offsetHeight;

      if (target.classList.contains('crop-handle')) {
        isResizing = true;
        resizeHandle = [...target.classList].find(c => ['nw','ne','sw','se'].includes(c));
      } else {
        isDragging = true;
      }

      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
    };

    const onMouseMove = (e) => {
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;

      if (isDragging) {
        overlay.style.left = (startLeft + dx) + 'px';
        overlay.style.top = (startTop + dy) + 'px';
      } else if (isResizing) {
        let newLeft = startLeft;
        let newTop = startTop;
        let newWidth = startWidth;
        let newHeight = startHeight;

        if (resizeHandle.includes('e')) newWidth = startWidth + dx;
        if (resizeHandle.includes('w')) { newWidth = startWidth - dx; newLeft = startLeft + dx; }
        if (resizeHandle.includes('s')) newHeight = startHeight + dy;
        if (resizeHandle.includes('n')) { newHeight = startHeight - dy; newTop = startTop + dy; }

        if (newWidth > 30 && newHeight > 30) {
          // Apply aspect ratio constraint if needed
          if (this.cropRatio !== 'free') {
            const [rw, rh] = this.cropRatio.split(':').map(Number);
            const ratio = rw / rh;
            if (resizeHandle.includes('e') || resizeHandle.includes('w')) {
              newHeight = newWidth / ratio;
            } else {
              newWidth = newHeight * ratio;
            }
          }

          overlay.style.left = newLeft + 'px';
          overlay.style.top = newTop + 'px';
          overlay.style.width = newWidth + 'px';
          overlay.style.height = newHeight + 'px';
        }
      }
    };

    const onMouseUp = () => {
      isDragging = false;
      isResizing = false;
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
    };

    overlay.addEventListener('mousedown', onMouseDown);

    // Touch support
    overlay.addEventListener('touchstart', (e) => {
      const touch = e.touches[0];
      onMouseDown({ preventDefault: () => e.preventDefault(), clientX: touch.clientX, clientY: touch.clientY, target: e.target });
    });
  }

  _updateCropOverlay() {
    if (!this.cropOverlay) return;
    // Recreate with current ratio
    const canvasRect = this.canvas.getBoundingClientRect();
    const canvasArea = document.querySelector('.editor-canvas-area');
    this._createCropOverlay();
  }

  async _applyCrop() {
    if (!this.cropOverlay) return;

    const canvasRect = this.canvas.getBoundingClientRect();
    const overlayRect = this.cropOverlay.getBoundingClientRect();

    // Calculate crop coordinates relative to actual image pixels
    const scaleX = this.canvas.width / canvasRect.width;
    const scaleY = this.canvas.height / canvasRect.height;

    const cropX = Math.max(0, (overlayRect.left - canvasRect.left) * scaleX);
    const cropY = Math.max(0, (overlayRect.top - canvasRect.top) * scaleY);
    const cropW = Math.min(this.canvas.width - cropX, overlayRect.width * scaleX);
    const cropH = Math.min(this.canvas.height - cropY, overlayRect.height * scaleY);

    if (cropW < 10 || cropH < 10) return;

    // Get cropped image data
    const imageData = this.ctx.getImageData(cropX, cropY, cropW, cropH);

    // Update original image to cropped version
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = cropW;
    tempCanvas.height = cropH;
    const tempCtx = tempCanvas.getContext('2d');
    tempCtx.putImageData(imageData, 0, 0);

    this.originalImage = await imageProcessor.loadImage(tempCanvas.toDataURL());
    this.rotation = 0;
    this.flipH = false;
    this.flipV = false;

    this._disableCropMode();
    this._renderPreview();
    this._enableCropMode();

    app.showToast('Crop applied');
  }
}

const editor = new PhotoEditor();
