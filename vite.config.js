import { defineConfig } from 'vite';
import devCerts from 'office-addin-dev-certs';

export default defineConfig(async ({ command }) => {
  const httpsOptions = await devCerts.getHttpsServerOptions();

  return {
    base: command === 'build' ? '/ToolExcel/' : '/',
    server: {
      host: 'localhost',
      port: 3000,
      strictPort: true,
      https: httpsOptions,
    },
    preview: {
      host: 'localhost',
      port: 4173,
      strictPort: true,
      https: httpsOptions,
    },
    build: {
      outDir: 'docs',
      sourcemap: true,
    },
  };
});
