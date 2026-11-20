const { createApp } = require('./app');

const app = createApp();
const { config, jobStore, localDb } = app.locals;

jobStore.startCleanupScheduler();

const server = app.listen(config.PORT, () => {
  console.log(`Kemenag Kota Metro Data Portal running at http://localhost:${config.PORT}`);
});

function shutdown(signal) {
  console.log(`\nReceived ${signal}, shutting down...`);
  jobStore.stopCleanupScheduler();
  server.close(() => {
    localDb.close();
    process.exit(0);
  });
}

process.on('SIGINT', () => shutdown('SIGINT'));
process.on('SIGTERM', () => shutdown('SIGTERM'));
