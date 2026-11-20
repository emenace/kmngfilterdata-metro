const { createApp } = require('./app');

const app = createApp();
const { config, jobStore } = app.locals;

jobStore.startCleanupScheduler();

const server = app.listen(config.PORT, () => {
  console.log(`Kemenag Kota Metro Excel Filter running at http://localhost:${config.PORT}`);
});

function shutdown(signal) {
  console.log(`\nReceived ${signal}, shutting down...`);
  jobStore.stopCleanupScheduler();
  server.close(() => {
    process.exit(0);
  });
}

process.on('SIGINT', () => shutdown('SIGINT'));
process.on('SIGTERM', () => shutdown('SIGTERM'));

