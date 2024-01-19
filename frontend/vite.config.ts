import {resolve} from "path";
import {defineConfig} from "vite";

// https://vitejs.dev/config/
export default defineConfig({
  root: "src",
  plugins: [],
  build: {
    minify: false,
    rollupOptions: {
      input: {
        main: resolve(__dirname, "src/index.html"),
        taskpane: resolve(__dirname, "src/taskpane/taskpane.html"),
      },
    },
  },
  esbuild: {
    minify: false,
    minifyIdentifiers: false,
    minifySyntax: false,
    minifyWhitespace: false,
  },
});
