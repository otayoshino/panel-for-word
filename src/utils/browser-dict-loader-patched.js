/**
 * src/utils/browser-dict-loader-patched.js
 *
 * BrowserDictionaryLoader のパッチ版。
 *
 * GitHub Pages (Fastly CDN) は .dat.gz ファイルを Content-Encoding: gzip 付きで
 * 配信することがある。その場合 XHR は自動展開済みの生バイトを返すため、
 * kuromoji 内の zlibjs が「二重展開」を試みて失敗し、コールバックが呼ばれなくなる。
 *
 * 対策: レスポンスの先頭 2 バイト (gzip マジックバイト 0x1f 0x8b) を確認し、
 *       - gzip のまま → DecompressionStream で展開してから返す
 *       - 既に展開済み → そのまま返す
 */
"use strict";

var DictionaryLoader = require("kuromoji/src/loader/DictionaryLoader");

function BrowserDictionaryLoader(dic_path) {
    DictionaryLoader.apply(this, [dic_path]);
}

BrowserDictionaryLoader.prototype = Object.create(DictionaryLoader.prototype);

BrowserDictionaryLoader.prototype.loadArrayBuffer = function (url, callback) {
    var xhr = new XMLHttpRequest();
    xhr.open("GET", url, true);
    xhr.responseType = "arraybuffer";
    xhr.onload = function () {
        if (this.status > 0 && this.status !== 200) {
            callback(xhr.statusText, null);
            return;
        }
        var data = new Uint8Array(this.response);

        // gzip マジックバイト (1f 8b) の有無で判別
        if (data.length >= 2 && data[0] === 0x1f && data[1] === 0x8b) {
            // まだ gzip 圧縮されている → DecompressionStream で展開
            try {
                var ds = new DecompressionStream("gzip");
                var writer = ds.writable.getWriter();
                writer.write(data);
                writer.close();
                new Response(ds.readable).arrayBuffer()
                    .then(function (buf) { callback(null, buf); })
                    .catch(function (err) { callback(err instanceof Error ? err.message : String(err), null); });
            } catch (e) {
                // DecompressionStream 未対応環境 → エラーとして返す
                callback("DecompressionStream unavailable: " + String(e), null);
            }
        } else {
            // Content-Encoding: gzip で既に展開済み → そのまま使用
            callback(null, this.response);
        }
    };
    xhr.onerror = function (err) {
        callback(err, null);
    };
    xhr.send();
};

module.exports = BrowserDictionaryLoader;
