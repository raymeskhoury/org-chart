import * as fs from "fs";
import {resolve} from "path";
import {defineConfig} from "vite";
import {viteStaticCopy} from "vite-plugin-static-copy";

let config = {
  root: "src",
  // resolve: {
  //   alias: [
  //     {
  //       find: /\/assets\/office-js\/(.+)/,
  //       replacement: `../node_modules/@microsoft/office-js/dist/$1`,
  //     },
  //     {
  //       find: /\/assets\/icons\/(.+)/,
  //       replacement: `../node_modules/@shoelace-style/shoelace/dist/assets/icons/$1`,
  //     },
  //   ],
  // },
  plugins: [
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
    https: {
      key: fs.readFileSync("./.cert/key.pem"),
      cert: fs.readFileSync("./.cert/cert.pem"),
    },
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
