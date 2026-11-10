// Bundles the library and the demo with esbuild. `--serve` starts a dev server.
import * as esbuild from 'esbuild';
import fs from 'node:fs';

const serve = process.argv.includes('--serve');
fs.mkdirSync('demo/dist', { recursive: true });
fs.copyFileSync('src/editor/editor.css', 'demo/dist/open-xlsx-editor.css');
const common = { bundle: true, format: 'esm', sourcemap: true, target: 'es2020', logLevel: 'info' };
await esbuild.build({ ...common, entryPoints: ['src/index.js'], outfile: 'dist/open-xlsx-editor.js', minify: true });
fs.copyFileSync('src/editor/editor.css', 'dist/open-xlsx-editor.css');
// (The co-editing module with Yjs, for the browser tests to load beside the demo.)
await esbuild.build({ ...common, entryPoints: ['demo/collab.js'], outfile: 'demo/dist/collab.js' });
const ctx = await esbuild.context({ ...common, entryPoints: ['demo/main.js'], outfile: 'demo/dist/demo.js' });
if (serve) {
  await ctx.watch();
  const { port } = await ctx.serve({ servedir: 'demo', port: 8080 });
  console.log(`demo: http://localhost:${port}/`);
} else {
  await ctx.rebuild();
  await ctx.dispose();
}
