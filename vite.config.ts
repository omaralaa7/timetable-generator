import { defineConfig } from 'vite';

// GitHub Pages serves this repository's root (Settings → Pages → "Deploy from a branch", main, /root),
// so the built page is committed there: `npm run build` writes index.html and assets/ next to the source.
// The page's own source lives in web/index.html.
export default defineConfig({
  root: 'web',
  base: './',
  server: { fs: { allow: ['..'] } },
  build: {
    outDir: '..',
    emptyOutDir: false,
    chunkSizeWarningLimit: 2500,
    rollupOptions: {
      output: { entryFileNames: 'assets/app.js', chunkFileNames: 'assets/[name].js', assetFileNames: 'assets/app[extname]' },
    },
  },
});
