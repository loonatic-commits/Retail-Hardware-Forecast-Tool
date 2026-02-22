/**
 * SnapVault - IndexedDB Database Layer
 * Handles all photo and album persistence using IndexedDB.
 */
class SnapVaultDB {
  constructor() {
    this.dbName = 'snapvault';
    this.dbVersion = 1;
    this.db = null;
  }

  async open() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.dbName, this.dbVersion);

      request.onupgradeneeded = (e) => {
        const db = e.target.result;

        // Photos store
        if (!db.objectStoreNames.contains('photos')) {
          const photoStore = db.createObjectStore('photos', { keyPath: 'id' });
          photoStore.createIndex('albumId', 'albumId', { unique: false });
          photoStore.createIndex('createdAt', 'createdAt', { unique: false });
          photoStore.createIndex('name', 'name', { unique: false });
        }

        // Albums store
        if (!db.objectStoreNames.contains('albums')) {
          const albumStore = db.createObjectStore('albums', { keyPath: 'id' });
          albumStore.createIndex('createdAt', 'createdAt', { unique: false });
        }
      };

      request.onsuccess = (e) => {
        this.db = e.target.result;
        resolve(this.db);
      };

      request.onerror = (e) => {
        reject(e.target.error);
      };
    });
  }

  // ---- Photo Operations ----

  async addPhoto(photo) {
    return this._transaction('photos', 'readwrite', (store) => {
      return store.add(photo);
    });
  }

  async getPhoto(id) {
    return this._transaction('photos', 'readonly', (store) => {
      return store.get(id);
    });
  }

  async getAllPhotos() {
    return this._transaction('photos', 'readonly', (store) => {
      return store.getAll();
    });
  }

  async getPhotosByAlbum(albumId) {
    return this._transaction('photos', 'readonly', (store) => {
      const index = store.index('albumId');
      return index.getAll(albumId);
    });
  }

  async updatePhoto(photo) {
    return this._transaction('photos', 'readwrite', (store) => {
      return store.put(photo);
    });
  }

  async deletePhoto(id) {
    return this._transaction('photos', 'readwrite', (store) => {
      return store.delete(id);
    });
  }

  async deletePhotos(ids) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('photos', 'readwrite');
      const store = tx.objectStore('photos');
      ids.forEach(id => store.delete(id));
      tx.oncomplete = () => resolve();
      tx.onerror = (e) => reject(e.target.error);
    });
  }

  // ---- Album Operations ----

  async addAlbum(album) {
    return this._transaction('albums', 'readwrite', (store) => {
      return store.add(album);
    });
  }

  async getAlbum(id) {
    return this._transaction('albums', 'readonly', (store) => {
      return store.get(id);
    });
  }

  async getAllAlbums() {
    return this._transaction('albums', 'readonly', (store) => {
      return store.getAll();
    });
  }

  async updateAlbum(album) {
    return this._transaction('albums', 'readwrite', (store) => {
      return store.put(album);
    });
  }

  async deleteAlbum(id) {
    return this._transaction('albums', 'readwrite', (store) => {
      return store.delete(id);
    });
  }

  // ---- Utility ----

  async getStorageEstimate() {
    const photos = await this.getAllPhotos();
    let totalBytes = 0;
    photos.forEach(p => {
      if (p.dataUrl) {
        totalBytes += p.dataUrl.length * 0.75; // Rough base64 to bytes
      }
      if (p.thumbnail) {
        totalBytes += p.thumbnail.length * 0.75;
      }
    });
    return {
      photoCount: photos.length,
      bytesUsed: totalBytes,
      mbUsed: (totalBytes / (1024 * 1024)).toFixed(1)
    };
  }

  async clearAll() {
    await this._transaction('photos', 'readwrite', (store) => store.clear());
    await this._transaction('albums', 'readwrite', (store) => store.clear());
  }

  async searchPhotos(query) {
    const photos = await this.getAllPhotos();
    const q = query.toLowerCase();
    return photos.filter(p =>
      (p.name && p.name.toLowerCase().includes(q)) ||
      (p.albumName && p.albumName.toLowerCase().includes(q))
    );
  }

  _transaction(storeName, mode, callback) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction(storeName, mode);
      const store = tx.objectStore(storeName);
      const request = callback(store);

      if (request && request.onsuccess !== undefined) {
        request.onsuccess = () => resolve(request.result);
        request.onerror = (e) => reject(e.target.error);
      } else {
        tx.oncomplete = () => resolve();
        tx.onerror = (e) => reject(e.target.error);
      }
    });
  }
}

// Singleton
const db = new SnapVaultDB();
