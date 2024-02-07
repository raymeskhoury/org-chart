import {resolve} from "path";
import {UserConfig, defineConfig} from "vite";
import mkcert from "vite-plugin-mkcert";
import {viteStaticCopy} from "vite-plugin-static-copy";

let config: UserConfig = {
  root: "src",
  plugins: [
    mkcert(),
    viteStaticCopy({
      targets: [
        {
          src: "../node_modules/@microsoft/office-js/dist/*",
          dest: "assets/office-js",
        },
        {
          src: "../node_modules/@shoelace-style/shoelace/dist/assets/icons/*",
          dest: "assets/icons/",
        },
      ],
    }),
  ],
  server: {
    port: 3000,
    https: {},
  },
  build: {
    rollupOptions: {
      input: {
        main: resolve(__dirname, "src/taskpane.html"),
      },
    },
    outDir: "../dist",
  },
};

// https://vitejs.dev/config/
export default defineConfig(config);
