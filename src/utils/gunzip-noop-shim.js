/**
 * src/utils/gunzip-noop-shim.js
 *
 * zlibjs/bin/gunzip.min.js の no-op 代替シム。
 * dict ファイルはビルド時に .dat.gz → .dat へ事前展開済みのため、
 * ブラウザ上での Gunzip は不要。受け取った Uint8Array をそのまま返す。
 */

function Gunzip(data) {
  this.data = data
}

Gunzip.prototype.decompress = function () {
  return this.data
}

module.exports = { Zlib: { Gunzip } }
