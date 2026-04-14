import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'fs'
import path from 'path'

const httpsConfig = (() => {
  try {
    return {
      key: fs.readFileSync('./certs/key.pem'),
      cert: fs.readFileSync('./certs/cert.pem'),
    }
  } catch {
    return undefined
  }
})()

// mkcert で生成した証明書を使用（Windows の信頼済みルート CA に登録済み）
// 証明書の生成: npx mkcert -install && npx mkcert -key-file certs/key.pem -cert-file certs/cert.pem localhost 127.0.0.1
export default defineConfig({
  plugins: [
    react(),
    {
      // kuromoji の BrowserDictionaryLoader を本番ビルドでパッチ版に差し替える。
      // alias・resolveId は @rollup/plugin-commonjs の仮想モジュール経由で importer が
      // kuromoji パスにならないため効かない。load フックでファイルパスを直接判定して差し替える。
      name: 'patch-kuromoji-browser-dict-loader',
      enforce: 'pre',
      load(id: string) {
        const normalized = id.replace(/\\/g, '/')
        if (normalized.includes('kuromoji') && normalized.endsWith('BrowserDictionaryLoader.js')) {
          return fs.readFileSync(path.resolve('./src/utils/browser-dict-loader-patched.js'), 'utf-8')
        }
        return null
      },
    },
    {
      // kuromoji 辞書ファイルを開発サーバーで /panel-for-word/dict/ として提供
      name: 'kuromoji-dict-serve',
      configureServer(server) {
        const dictDir = path.resolve('./node_modules/kuromoji/dict')
        server.middlewares.use('/panel-for-word/dict/', (req, res, next) => {
          const fileName = (req.url ?? '').replace(/^\/?/, '')
          if (!fileName) { next(); return }
          const filePath = path.join(dictDir, fileName)
          try {
            const buf = fs.readFileSync(filePath)
            res.setHeader('Content-Type', 'application/octet-stream')
            res.end(buf)
          } catch {
            next()
          }
        })
      },
    },
  ],
  base: '/panel-for-word/',
  resolve: {
    alias: {
      // kuromoji の DictionaryLoader が require("path") を使うため、
      // ブラウザ環境で動く最小シムに差し替える（path-browserify は https:// を壊すため不可）
      path: path.resolve('./src/utils/path-browser-shim.js'),
      // GitHub Pages が .dat.gz を Content-Encoding: gzip で配信し XHR が自動展開する問題を吸収。
      // gzip マジックバイト判別で展開済み/未展開を自動判定するパッチ版ローダーに差し替える。
      'kuromoji/src/loader/BrowserDictionaryLoader': path.resolve('./src/utils/browser-dict-loader-patched.js'),
    },
  },
  server: {
    port: 3000,
    strictPort: true,
    https: httpsConfig,
  },
  build: {
    outDir: 'dist',
    commonjsOptions: {
      // kuromoji は CJS パッケージのため本番ビルドでも確実に変換する
      include: [/kuromoji/, /node_modules/],
    },
  },
  optimizeDeps: {
    include: ['kuromoji'],
  },
})
