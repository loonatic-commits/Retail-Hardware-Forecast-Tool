/**
 * SnapVault - Main Application Controller
 * Manages views, albums, photos, camera scanning, viewer, multi-select, export, search.
 */
class SnapVaultApp {
  constructor() {
    this.currentView = 'albums';
    this.currentAlbumId = null;
    this.selectMode = false;
    this.selectedPhotos = new Set();
    this.allPhotos = [];
    this.allAlbums = [];
    this.viewerPhotos = [];
    this.viewerIndex = 0;
    this.cameraStream = null;
    this.facingMode = 'environment';
  }

  async init() {
    await db.open();
    this._bindEvents();
    await this._loadData();
    this._renderAlbums();
    this._renderAllPhotos();
    this._updateStorageInfo();
  }

  // ========================================
  // Event Bindings
  // ========================================

  _bindEvents() {
    // Navigation
    document.querySelectorAll('.nav-btn').forEach(btn => {
      btn.addEventListener('click', () => this.switchView(btn.dataset.view));
    });

    // Upload
    document.getElementById('upload-btn').addEventListener('click', () => {
      document.getElementById('file-input').click();
    });

    document.getElementById('file-input').addEventListener('change', (e) => {
      this._handleFileUpload(e.target.files);
      e.target.value = '';
    });

    // Sidebar
    document.getElementById('sidebar-toggle').addEventListener('click', () => {
      document.getElementById('sidebar').classList.toggle('hidden');
    });

    document.getElementById('new-album-btn').addEventListener('click', () => this.createAlbumPrompt());
    document.getElementById('create-album-btn').addEventListener('click', () => this.createAlbumPrompt());

    // Sidebar footer
    document.getElementById('export-all-btn').addEventListener('click', () => this._exportAll());
    document.getElementById('clear-data-btn').addEventListener('click', () => this._clearAllData());

    // Search
    document.getElementById('search-btn').addEventListener('click', () => {
      const bar = document.getElementById('search-bar');
      bar.classList.toggle('hidden');
      if (!bar.classList.contains('hidden')) {
        document.getElementById('search-input').focus();
      }
    });
    document.getElementById('search-close').addEventListener('click', () => {
      document.getElementById('search-bar').classList.add('hidden');
      document.getElementById('search-input').value = '';
      this._renderAllPhotos();
    });
    document.getElementById('search-input').addEventListener('input', (e) => {
      this._handleSearch(e.target.value);
    });

    // Sort
    document.getElementById('sort-select').addEventListener('change', () => {
      this._renderAllPhotos();
    });

    // Select mode
    document.getElementById('select-mode-btn').addEventListener('click', () => {
      this.toggleSelectMode();
    });
    document.getElementById('sel-cancel').addEventListener('click', () => {
      this.toggleSelectMode(false);
    });
    document.getElementById('sel-delete').addEventListener('click', () => this._deleteSelected());
    document.getElementById('sel-download').addEventListener('click', () => this._downloadSelected());
    document.getElementById('sel-add-album').addEventListener('click', () => this._addSelectedToAlbum());

    // Album detail
    document.getElementById('album-back-btn').addEventListener('click', () => this.switchView('albums'));
    document.getElementById('album-add-photos').addEventListener('click', () => this._addPhotosToCurrentAlbum());
    document.getElementById('album-edit-btn').addEventListener('click', () => this._editAlbum());
    document.getElementById('album-delete-btn').addEventListener('click', () => this._deleteAlbum());
    document.getElementById('album-empty-add').addEventListener('click', () => this._addPhotosToCurrentAlbum());

    // Viewer
    document.getElementById('viewer-close').addEventListener('click', () => this.closeViewer());
    document.getElementById('viewer-prev').addEventListener('click', () => this.viewerNav(-1));
    document.getElementById('viewer-next').addEventListener('click', () => this.viewerNav(1));
    document.getElementById('viewer-edit').addEventListener('click', () => this._editCurrentViewerPhoto());
    document.getElementById('viewer-info').addEventListener('click', () => this._toggleInfoPanel());
    document.getElementById('viewer-download').addEventListener('click', () => this._downloadCurrentPhoto());
    document.getElementById('viewer-delete').addEventListener('click', () => this._deleteCurrentViewerPhoto());

    // Keyboard navigation
    document.addEventListener('keydown', (e) => {
      const viewer = document.getElementById('photo-viewer');
      if (!viewer.classList.contains('hidden')) {
        if (e.key === 'ArrowLeft') this.viewerNav(-1);
        else if (e.key === 'ArrowRight') this.viewerNav(1);
        else if (e.key === 'Escape') this.closeViewer();
      }
      const editorEl = document.getElementById('photo-editor');
      if (!editorEl.classList.contains('hidden') && e.key === 'Escape') {
        editor.close();
      }
    });

    // Scan
    document.getElementById('scan-capture-btn').addEventListener('click', () => this._capturePhoto());
    document.getElementById('scan-switch-cam').addEventListener('click', () => this._switchCamera());
    document.getElementById('scan-retake').addEventListener('click', () => this._retakeScan());
    document.getElementById('scan-enhance').addEventListener('click', () => this._enhanceScan());
    document.getElementById('scan-save').addEventListener('click', () => this._saveScan());

    // Close sidebar on outside click
    document.addEventListener('click', (e) => {
      const sidebar = document.getElementById('sidebar');
      const toggle = document.getElementById('sidebar-toggle');
      if (!sidebar.classList.contains('hidden') &&
          !sidebar.contains(e.target) &&
          !toggle.contains(e.target)) {
        sidebar.classList.add('hidden');
      }
    });

    // Modal close on overlay click
    document.getElementById('modal-overlay').addEventListener('click', (e) => {
      if (e.target === e.currentTarget) this.closeModal();
    });

    // Drag and drop
    const main = document.getElementById('main-content');
    main.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); });
    main.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (e.dataTransfer.files.length > 0) {
        this._handleFileUpload(e.dataTransfer.files);
      }
    });
  }

  // ========================================
  // Data Loading
  // ========================================

  async _loadData() {
    this.allPhotos = await db.getAllPhotos();
    this.allAlbums = await db.getAllAlbums();
  }

  // ========================================
  // View Management
  // ========================================

  switchView(view) {
    this.currentView = view;
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

    const navBtn = document.querySelector(`.nav-btn[data-view="${view}"]`);
    if (navBtn) navBtn.classList.add('active');

    if (view === 'albums') {
      document.getElementById('albums-view').classList.add('active');
      this._renderAlbums();
    } else if (view === 'all-photos') {
      document.getElementById('all-photos-view').classList.add('active');
      this._renderAllPhotos();
    } else if (view === 'scan') {
      document.getElementById('scan-view').classList.add('active');
      this._startCamera();
    } else if (view === 'album-detail') {
      document.getElementById('album-detail-view').classList.add('active');
      this._renderAlbumDetail();
    }

    // Stop camera when leaving scan view
    if (view !== 'scan' && this.cameraStream) {
      this._stopCamera();
    }
  }

  async refreshCurrentView() {
    await this._loadData();
    if (this.currentView === 'albums') this._renderAlbums();
    else if (this.currentView === 'all-photos') this._renderAllPhotos();
    else if (this.currentView === 'album-detail') this._renderAlbumDetail();
    this._updateSidebar();
    this._updateStorageInfo();
  }

  // ========================================
  // Albums
  // ========================================

  _renderAlbums() {
    const grid = document.getElementById('albums-grid');
    const empty = document.getElementById('albums-empty');

    if (this.allAlbums.length === 0) {
      grid.innerHTML = '';
      empty.classList.remove('hidden');
      return;
    }

    empty.classList.add('hidden');
    grid.innerHTML = '';

    this.allAlbums.forEach(album => {
      const albumPhotos = this.allPhotos.filter(p => p.albumId === album.id);
      const card = document.createElement('div');
      card.className = 'album-card';
      card.addEventListener('click', () => {
        this.currentAlbumId = album.id;
        this.switchView('album-detail');
      });

      // Cover: show up to 4 thumbnails
      const cover = document.createElement('div');
      cover.className = 'album-card-cover';
      for (let i = 0; i < 4; i++) {
        if (albumPhotos[i]) {
          const img = document.createElement('img');
          img.src = albumPhotos[i].thumbnail || albumPhotos[i].dataUrl;
          img.alt = albumPhotos[i].name;
          img.loading = 'lazy';
          cover.appendChild(img);
        } else {
          const ph = document.createElement('div');
          ph.className = 'placeholder';
          cover.appendChild(ph);
        }
      }

      const info = document.createElement('div');
      info.className = 'album-card-info';
      info.innerHTML = `<h3>${this._escapeHtml(album.name)}</h3><span>${albumPhotos.length} photo${albumPhotos.length !== 1 ? 's' : ''}</span>`;

      card.appendChild(cover);
      card.appendChild(info);
      grid.appendChild(card);
    });
  }

  async createAlbumPrompt() {
    this.showModal(`
      <h3>Create New Album</h3>
      <input type="text" id="album-name-input" placeholder="Album name" autofocus>
      <div class="modal-actions">
        <button class="small-btn" onclick="app.closeModal()">Cancel</button>
        <button class="small-btn primary" onclick="app.createAlbum()">Create</button>
      </div>
    `);
    setTimeout(() => {
      const input = document.getElementById('album-name-input');
      if (input) {
        input.focus();
        input.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') app.createAlbum();
        });
      }
    }, 100);
  }

  async createAlbum() {
    const input = document.getElementById('album-name-input');
    const name = input ? input.value.trim() : '';
    if (!name) {
      this.showToast('Please enter an album name', 'error');
      return;
    }

    const album = {
      id: this._generateId(),
      name,
      createdAt: Date.now(),
    };

    await db.addAlbum(album);
    this.closeModal();
    await this.refreshCurrentView();
    this.showToast(`Album "${name}" created`);
  }

  _renderAlbumDetail() {
    const album = this.allAlbums.find(a => a.id === this.currentAlbumId);
    if (!album) {
      this.switchView('albums');
      return;
    }

    document.getElementById('album-detail-title').textContent = album.name;
    const photos = this.allPhotos.filter(p => p.albumId === album.id);
    const grid = document.getElementById('album-photos-grid');
    const empty = document.getElementById('album-photos-empty');

    if (photos.length === 0) {
      grid.innerHTML = '';
      empty.classList.remove('hidden');
      return;
    }

    empty.classList.add('hidden');
    grid.innerHTML = '';
    this._renderPhotoGrid(grid, photos);
  }

  async _editAlbum() {
    const album = this.allAlbums.find(a => a.id === this.currentAlbumId);
    if (!album) return;

    this.showModal(`
      <h3>Rename Album</h3>
      <input type="text" id="album-rename-input" value="${this._escapeHtml(album.name)}" autofocus>
      <div class="modal-actions">
        <button class="small-btn" onclick="app.closeModal()">Cancel</button>
        <button class="small-btn primary" onclick="app.renameAlbum()">Save</button>
      </div>
    `);
    setTimeout(() => {
      const input = document.getElementById('album-rename-input');
      if (input) {
        input.focus();
        input.select();
        input.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') app.renameAlbum();
        });
      }
    }, 100);
  }

  async renameAlbum() {
    const input = document.getElementById('album-rename-input');
    const name = input ? input.value.trim() : '';
    if (!name) return;

    const album = this.allAlbums.find(a => a.id === this.currentAlbumId);
    if (!album) return;

    album.name = name;
    await db.updateAlbum(album);

    // Update album name on photos
    const photos = this.allPhotos.filter(p => p.albumId === album.id);
    for (const photo of photos) {
      photo.albumName = name;
      await db.updatePhoto(photo);
    }

    this.closeModal();
    await this.refreshCurrentView();
    this.showToast('Album renamed');
  }

  async _deleteAlbum() {
    const album = this.allAlbums.find(a => a.id === this.currentAlbumId);
    if (!album) return;

    this.showModal(`
      <h3>Delete Album</h3>
      <p>Delete "${this._escapeHtml(album.name)}"? Photos will be moved out of this album but not deleted.</p>
      <div class="modal-actions">
        <button class="small-btn" onclick="app.closeModal()">Cancel</button>
        <button class="small-btn danger" onclick="app.confirmDeleteAlbum()">Delete</button>
      </div>
    `);
  }

  async confirmDeleteAlbum() {
    // Remove album association from photos
    const photos = this.allPhotos.filter(p => p.albumId === this.currentAlbumId);
    for (const photo of photos) {
      photo.albumId = null;
      photo.albumName = null;
      await db.updatePhoto(photo);
    }

    await db.deleteAlbum(this.currentAlbumId);
    this.closeModal();
    this.currentAlbumId = null;
    await this.refreshCurrentView();
    this.switchView('albums');
    this.showToast('Album deleted');
  }

  async _addPhotosToCurrentAlbum() {
    // Show picker with unassigned photos
    const unassigned = this.allPhotos.filter(p => !p.albumId || p.albumId !== this.currentAlbumId);
    if (unassigned.length === 0) {
      this.showToast('No photos available to add', 'error');
      return;
    }

    let html = '<h3>Add Photos to Album</h3><div class="photos-grid" style="max-height:300px;overflow-y:auto;margin-bottom:16px;">';
    unassigned.forEach(photo => {
      html += `<div class="photo-card" data-id="${photo.id}" onclick="this.classList.toggle('selected')">
        <img src="${photo.thumbnail || photo.dataUrl}" alt="${this._escapeHtml(photo.name)}" loading="lazy">
        <div class="select-check"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg></div>
      </div>`;
    });
    html += '</div><div class="modal-actions"><button class="small-btn" onclick="app.closeModal()">Cancel</button><button class="small-btn primary" onclick="app.confirmAddToAlbum()">Add Selected</button></div>';

    this.showModal(html);
  }

  async confirmAddToAlbum() {
    const selected = document.querySelectorAll('#modal-content .photo-card.selected');
    const album = this.allAlbums.find(a => a.id === this.currentAlbumId);
    if (!album || selected.length === 0) {
      this.showToast('No photos selected', 'error');
      return;
    }

    for (const card of selected) {
      const photo = this.allPhotos.find(p => p.id === card.dataset.id);
      if (photo) {
        photo.albumId = album.id;
        photo.albumName = album.name;
        await db.updatePhoto(photo);
      }
    }

    this.closeModal();
    await this.refreshCurrentView();
    this.showToast(`${selected.length} photo(s) added to album`);
  }

  // ========================================
  // Photos
  // ========================================

  _renderAllPhotos() {
    const grid = document.getElementById('photos-grid');
    const empty = document.getElementById('photos-empty');

    let photos = [...this.allPhotos];

    // Sort
    const sort = document.getElementById('sort-select').value;
    if (sort === 'newest') photos.sort((a, b) => b.createdAt - a.createdAt);
    else if (sort === 'oldest') photos.sort((a, b) => a.createdAt - b.createdAt);
    else if (sort === 'name') photos.sort((a, b) => (a.name || '').localeCompare(b.name || ''));

    if (photos.length === 0) {
      grid.innerHTML = '';
      empty.classList.remove('hidden');
      return;
    }

    empty.classList.add('hidden');
    grid.innerHTML = '';
    this._renderPhotoGrid(grid, photos);
  }

  _renderPhotoGrid(grid, photos) {
    if (this.selectMode) grid.classList.add('select-mode');
    else grid.classList.remove('select-mode');

    photos.forEach((photo, index) => {
      const card = document.createElement('div');
      card.className = 'photo-card' + (this.selectedPhotos.has(photo.id) ? ' selected' : '');
      card.dataset.id = photo.id;

      const img = document.createElement('img');
      img.src = photo.thumbnail || photo.dataUrl;
      img.alt = photo.name || 'Photo';
      img.loading = 'lazy';

      const overlay = document.createElement('div');
      overlay.className = 'photo-overlay';
      const dateStr = new Date(photo.createdAt).toLocaleDateString();
      overlay.innerHTML = `<div><span class="photo-name">${this._escapeHtml(photo.name || 'Untitled')}</span><span class="photo-date">${dateStr}</span></div>`;

      const check = document.createElement('div');
      check.className = 'select-check';
      check.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>';

      card.appendChild(img);
      card.appendChild(overlay);
      card.appendChild(check);

      card.addEventListener('click', () => {
        if (this.selectMode) {
          this._togglePhotoSelection(photo.id, card);
        } else {
          this.openViewer(photos, index);
        }
      });

      grid.appendChild(card);
    });
  }

  async _handleFileUpload(files) {
    if (!files || files.length === 0) return;

    const fileArray = Array.from(files).filter(f => f.type.startsWith('image/'));
    if (fileArray.length === 0) {
      this.showToast('No image files selected', 'error');
      return;
    }

    this.showToast(`Importing ${fileArray.length} photo(s)...`);

    for (const file of fileArray) {
      try {
        const dataUrl = await imageProcessor.readFileAsDataURL(file);
        const thumbnail = await imageProcessor.generateThumbnail(dataUrl);

        const photo = {
          id: this._generateId(),
          name: file.name.replace(/\.[^.]+$/, ''),
          dataUrl,
          thumbnail,
          albumId: this.currentAlbumId || null,
          albumName: this.currentAlbumId ? (this.allAlbums.find(a => a.id === this.currentAlbumId)?.name || null) : null,
          createdAt: Date.now(),
          fileSize: file.size,
          fileType: file.type,
          width: 0,
          height: 0,
        };

        // Get dimensions
        const img = await imageProcessor.loadImage(dataUrl);
        photo.width = img.width;
        photo.height = img.height;

        await db.addPhoto(photo);
      } catch (err) {
        console.error('Failed to import file:', file.name, err);
      }
    }

    await this.refreshCurrentView();
    this.showToast(`${fileArray.length} photo(s) imported`, 'success');
  }

  async _handleSearch(query) {
    if (!query.trim()) {
      this._renderAllPhotos();
      return;
    }

    const results = await db.searchPhotos(query);
    const grid = document.getElementById('photos-grid');
    const empty = document.getElementById('photos-empty');

    if (results.length === 0) {
      grid.innerHTML = '<div class="empty-state"><p>No photos match your search</p></div>';
      empty.classList.add('hidden');
      return;
    }

    empty.classList.add('hidden');
    grid.innerHTML = '';
    this._renderPhotoGrid(grid, results);
  }

  // ========================================
  // Selection Mode
  // ========================================

  toggleSelectMode(force) {
    this.selectMode = force !== undefined ? force : !this.selectMode;
    this.selectedPhotos.clear();

    const toolbar = document.getElementById('selection-toolbar');
    const btn = document.getElementById('select-mode-btn');

    if (this.selectMode) {
      toolbar.classList.remove('hidden');
      btn.textContent = 'Cancel';
    } else {
      toolbar.classList.add('hidden');
      btn.textContent = 'Select';
    }

    this._updateSelectionCount();
    this.refreshCurrentView();
  }

  _togglePhotoSelection(photoId, card) {
    if (this.selectedPhotos.has(photoId)) {
      this.selectedPhotos.delete(photoId);
      card.classList.remove('selected');
    } else {
      this.selectedPhotos.add(photoId);
      card.classList.add('selected');
    }
    this._updateSelectionCount();
  }

  _updateSelectionCount() {
    document.getElementById('selection-count').textContent =
      `${this.selectedPhotos.size} selected`;
  }

  async _deleteSelected() {
    if (this.selectedPhotos.size === 0) return;

    const count = this.selectedPhotos.size;
    this.showModal(`
      <h3>Delete Photos</h3>
      <p>Delete ${count} selected photo(s)? This cannot be undone.</p>
      <div class="modal-actions">
        <button class="small-btn" onclick="app.closeModal()">Cancel</button>
        <button class="small-btn danger" onclick="app.confirmDeleteSelected()">Delete</button>
      </div>
    `);
  }

  async confirmDeleteSelected() {
    await db.deletePhotos([...this.selectedPhotos]);
    this.closeModal();
    this.toggleSelectMode(false);
    await this.refreshCurrentView();
    this.showToast('Photos deleted', 'success');
  }

  async _downloadSelected() {
    for (const id of this.selectedPhotos) {
      const photo = this.allPhotos.find(p => p.id === id);
      if (photo) this._downloadPhoto(photo);
    }
    this.showToast(`${this.selectedPhotos.size} photo(s) downloading`);
  }

  async _addSelectedToAlbum() {
    if (this.allAlbums.length === 0) {
      this.showToast('Create an album first', 'error');
      return;
    }

    let html = '<h3>Add to Album</h3><div class="album-picker">';
    this.allAlbums.forEach(album => {
      html += `<div class="album-picker-item" data-id="${album.id}" onclick="this.classList.toggle('selected');document.querySelectorAll('.album-picker-item').forEach(i=>{ if(i!==this) i.classList.remove('selected') })">
        <strong>${this._escapeHtml(album.name)}</strong>
      </div>`;
    });
    html += '</div><div class="modal-actions"><button class="small-btn" onclick="app.closeModal()">Cancel</button><button class="small-btn primary" onclick="app.confirmAddSelectedToAlbum()">Add</button></div>';

    this.showModal(html);
  }

  async confirmAddSelectedToAlbum() {
    const selected = document.querySelector('.album-picker-item.selected');
    if (!selected) {
      this.showToast('Select an album', 'error');
      return;
    }

    const albumId = selected.dataset.id;
    const album = this.allAlbums.find(a => a.id === albumId);

    for (const photoId of this.selectedPhotos) {
      const photo = this.allPhotos.find(p => p.id === photoId);
      if (photo) {
        photo.albumId = albumId;
        photo.albumName = album.name;
        await db.updatePhoto(photo);
      }
    }

    this.closeModal();
    this.toggleSelectMode(false);
    await this.refreshCurrentView();
    this.showToast(`Photos added to "${album.name}"`, 'success');
  }

  // ========================================
  // Photo Viewer / Lightbox
  // ========================================

  openViewer(photos, index) {
    this.viewerPhotos = photos;
    this.viewerIndex = index;

    document.getElementById('photo-viewer').classList.remove('hidden');
    document.getElementById('info-panel').classList.add('hidden');
    this._updateViewerImage();
  }

  closeViewer() {
    document.getElementById('photo-viewer').classList.add('hidden');
    document.getElementById('info-panel').classList.add('hidden');
  }

  viewerNav(direction) {
    this.viewerIndex += direction;
    if (this.viewerIndex < 0) this.viewerIndex = this.viewerPhotos.length - 1;
    if (this.viewerIndex >= this.viewerPhotos.length) this.viewerIndex = 0;
    this._updateViewerImage();
  }

  _updateViewerImage() {
    const photo = this.viewerPhotos[this.viewerIndex];
    if (!photo) return;

    document.getElementById('viewer-image').src = photo.dataUrl;
    document.getElementById('viewer-title').textContent = photo.name || 'Untitled';

    // Update info panel if open
    if (!document.getElementById('info-panel').classList.contains('hidden')) {
      this._renderInfoPanel(photo);
    }
  }

  _toggleInfoPanel() {
    const panel = document.getElementById('info-panel');
    panel.classList.toggle('hidden');
    if (!panel.classList.contains('hidden')) {
      const photo = this.viewerPhotos[this.viewerIndex];
      this._renderInfoPanel(photo);
    }
  }

  _renderInfoPanel(photo) {
    const content = document.getElementById('info-content');
    const date = new Date(photo.createdAt);
    const editDate = photo.editedAt ? new Date(photo.editedAt).toLocaleString() : 'Never';
    const fileSize = photo.fileSize ? this._formatBytes(photo.fileSize) : 'Unknown';

    content.innerHTML = `
      <input class="info-name-input" value="${this._escapeHtml(photo.name || '')}" placeholder="Photo name" onchange="app.renamePhoto('${photo.id}', this.value)">
      <div class="info-row"><span class="label">Date</span><span class="value">${date.toLocaleString()}</span></div>
      <div class="info-row"><span class="label">Dimensions</span><span class="value">${photo.width || '?'} x ${photo.height || '?'}</span></div>
      <div class="info-row"><span class="label">File Size</span><span class="value">${fileSize}</span></div>
      <div class="info-row"><span class="label">Format</span><span class="value">${photo.fileType || 'image/jpeg'}</span></div>
      <div class="info-row"><span class="label">Album</span><span class="value">${photo.albumName || 'None'}</span></div>
      <div class="info-row"><span class="label">Last Edited</span><span class="value">${editDate}</span></div>
    `;
  }

  async renamePhoto(id, newName) {
    const photo = this.allPhotos.find(p => p.id === id);
    if (!photo) return;
    photo.name = newName.trim();
    await db.updatePhoto(photo);
    this._updateViewerImage();
    this.showToast('Photo renamed');
  }

  async _editCurrentViewerPhoto() {
    const photo = this.viewerPhotos[this.viewerIndex];
    if (!photo) return;
    this.closeViewer();
    editor.open(photo);
  }

  _downloadCurrentPhoto() {
    const photo = this.viewerPhotos[this.viewerIndex];
    if (photo) this._downloadPhoto(photo);
  }

  async _deleteCurrentViewerPhoto() {
    const photo = this.viewerPhotos[this.viewerIndex];
    if (!photo) return;

    this.showModal(`
      <h3>Delete Photo</h3>
      <p>Delete "${this._escapeHtml(photo.name || 'this photo')}"? This cannot be undone.</p>
      <div class="modal-actions">
        <button class="small-btn" onclick="app.closeModal()">Cancel</button>
        <button class="small-btn danger" onclick="app.confirmDeleteViewerPhoto()">Delete</button>
      </div>
    `);
  }

  async confirmDeleteViewerPhoto() {
    const photo = this.viewerPhotos[this.viewerIndex];
    await db.deletePhoto(photo.id);
    this.viewerPhotos.splice(this.viewerIndex, 1);

    this.closeModal();

    if (this.viewerPhotos.length === 0) {
      this.closeViewer();
    } else {
      if (this.viewerIndex >= this.viewerPhotos.length) this.viewerIndex = 0;
      this._updateViewerImage();
    }

    await this.refreshCurrentView();
    this.showToast('Photo deleted');
  }

  // ========================================
  // Camera / Scan
  // ========================================

  async _startCamera() {
    try {
      const constraints = {
        video: {
          facingMode: this.facingMode,
          width: { ideal: 1920 },
          height: { ideal: 1080 },
        },
      };

      this.cameraStream = await navigator.mediaDevices.getUserMedia(constraints);
      document.getElementById('camera-feed').srcObject = this.cameraStream;
      document.getElementById('scan-preview').classList.add('hidden');
      document.getElementById('camera-container').style.display = 'block';
    } catch (err) {
      console.error('Camera access failed:', err);
      this.showToast('Camera access denied. You can still upload photos.', 'error');
    }
  }

  _stopCamera() {
    if (this.cameraStream) {
      this.cameraStream.getTracks().forEach(t => t.stop());
      this.cameraStream = null;
    }
  }

  async _switchCamera() {
    this.facingMode = this.facingMode === 'environment' ? 'user' : 'environment';
    this._stopCamera();
    await this._startCamera();
  }

  _capturePhoto() {
    const video = document.getElementById('camera-feed');
    if (!video.srcObject) return;

    const canvas = document.createElement('canvas');
    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(video, 0, 0);

    const dataUrl = canvas.toDataURL('image/jpeg', 0.92);
    document.getElementById('scan-result').src = dataUrl;
    document.getElementById('scan-result').dataset.dataUrl = dataUrl;
    document.getElementById('scan-preview').classList.remove('hidden');
    document.getElementById('camera-container').style.display = 'none';
  }

  _retakeScan() {
    document.getElementById('scan-preview').classList.add('hidden');
    document.getElementById('camera-container').style.display = 'block';
  }

  async _enhanceScan() {
    const img = document.getElementById('scan-result');
    const dataUrl = img.dataset.dataUrl || img.src;

    const loadedImg = await imageProcessor.loadImage(dataUrl);
    const canvas = document.createElement('canvas');
    canvas.width = loadedImg.width;
    canvas.height = loadedImg.height;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(loadedImg, 0, 0);

    const adjustments = imageProcessor.autoEnhance(ctx, canvas.width, canvas.height);
    // Re-draw and apply
    ctx.drawImage(loadedImg, 0, 0);
    imageProcessor.applyAdjustments(ctx, canvas.width, canvas.height, adjustments);

    const enhanced = canvas.toDataURL('image/jpeg', 0.92);
    img.src = enhanced;
    img.dataset.dataUrl = enhanced;
    this.showToast('Photo enhanced');
  }

  async _saveScan() {
    const img = document.getElementById('scan-result');
    const dataUrl = img.dataset.dataUrl || img.src;

    const thumbnail = await imageProcessor.generateThumbnail(dataUrl);
    const loadedImg = await imageProcessor.loadImage(dataUrl);

    const photo = {
      id: this._generateId(),
      name: `Scan ${new Date().toLocaleString()}`,
      dataUrl,
      thumbnail,
      albumId: null,
      albumName: null,
      createdAt: Date.now(),
      fileSize: Math.round(dataUrl.length * 0.75),
      fileType: 'image/jpeg',
      width: loadedImg.width,
      height: loadedImg.height,
    };

    await db.addPhoto(photo);
    await this.refreshCurrentView();
    this.showToast('Photo saved', 'success');

    // Go back to camera
    this._retakeScan();
  }

  // ========================================
  // Export / Download
  // ========================================

  _downloadPhoto(photo) {
    const link = document.createElement('a');
    link.href = photo.dataUrl;
    link.download = (photo.name || 'photo') + '.jpg';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  async _exportAll() {
    if (this.allPhotos.length === 0) {
      this.showToast('No photos to export', 'error');
      return;
    }

    this.showToast(`Downloading ${this.allPhotos.length} photos...`);

    // Download each photo individually (no zip library needed)
    for (let i = 0; i < this.allPhotos.length; i++) {
      setTimeout(() => {
        this._downloadPhoto(this.allPhotos[i]);
      }, i * 500); // Stagger downloads
    }
  }

  async _clearAllData() {
    this.showModal(`
      <h3>Clear All Data</h3>
      <p>This will permanently delete all photos and albums. This cannot be undone.</p>
      <div class="modal-actions">
        <button class="small-btn" onclick="app.closeModal()">Cancel</button>
        <button class="small-btn danger" onclick="app.confirmClearAll()">Clear Everything</button>
      </div>
    `);
  }

  async confirmClearAll() {
    await db.clearAll();
    this.closeModal();
    await this.refreshCurrentView();
    this.showToast('All data cleared');
  }

  // ========================================
  // Sidebar
  // ========================================

  _updateSidebar() {
    const list = document.getElementById('album-list');
    list.innerHTML = '';

    this.allAlbums.forEach(album => {
      const photos = this.allPhotos.filter(p => p.albumId === album.id);
      const li = document.createElement('li');

      const thumb = photos[0]
        ? `<img class="album-thumb" src="${photos[0].thumbnail || photos[0].dataUrl}" alt="" loading="lazy">`
        : '<div class="album-thumb" style="display:flex;align-items:center;justify-content:center;background:var(--bg);border-radius:6px;">&#128247;</div>';

      li.innerHTML = `${thumb}<span>${this._escapeHtml(album.name)}</span><span class="album-count">${photos.length}</span>`;
      li.addEventListener('click', () => {
        this.currentAlbumId = album.id;
        this.switchView('album-detail');
        document.getElementById('sidebar').classList.add('hidden');
      });
      list.appendChild(li);
    });
  }

  async _updateStorageInfo() {
    const info = await db.getStorageEstimate();
    document.getElementById('photo-count').textContent = `${info.photoCount} photo${info.photoCount !== 1 ? 's' : ''}`;
    document.getElementById('storage-used').textContent = `${info.mbUsed} MB used`;
  }

  // ========================================
  // Modal & Toast
  // ========================================

  showModal(html) {
    document.getElementById('modal-content').innerHTML = html;
    document.getElementById('modal-overlay').classList.remove('hidden');
  }

  closeModal() {
    document.getElementById('modal-overlay').classList.add('hidden');
  }

  showToast(message, type = '') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = 'toast' + (type ? ` ${type}` : '');
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(20px)';
      toast.style.transition = '0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  }

  // ========================================
  // Utilities
  // ========================================

  _generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 9);
  }

  _escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  _formatBytes(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
  }
}

// ---- Boot ----
const app = new SnapVaultApp();
document.addEventListener('DOMContentLoaded', () => app.init());
