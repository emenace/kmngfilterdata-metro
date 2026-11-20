const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Manages temporary per-job export directories and TTL-based cleanup.
 */
class JobStore {
  constructor({ tmpDir, ttlMs }) {
    this.tmpDir = path.resolve(tmpDir);
    this.ttlMs = ttlMs;
    this.jobs = new Map();
    this.cleanupTimer = null;
    fs.mkdirSync(this.tmpDir, { recursive: true });
    this.pruneOrphanDirs();
  }

  isValidJobId(jobId) {
    return typeof jobId === 'string' && UUID_REGEX.test(jobId);
  }

  createJobDir() {
    const jobId = crypto.randomUUID();
    const jobDir = path.join(this.tmpDir, jobId);
    fs.mkdirSync(jobDir, { recursive: true });
    return { jobId, jobDir };
  }

  saveJob(jobId, metadata) {
    this.jobs.set(jobId, {
      ...metadata,
      createdAt: Date.now(),
    });
  }

  removeJobDir(jobId) {
    this.jobs.delete(jobId);
    if (this.isValidJobId(jobId)) {
      const jobDir = path.join(this.tmpDir, jobId);
      fs.rmSync(jobDir, { recursive: true, force: true });
    }
  }

  getJobFile(jobId, format) {
    if (!this.isValidJobId(jobId)) {
      return null;
    }

    const job = this.jobs.get(jobId);
    if (!job) {
      return null;
    }

    if (Date.now() - job.createdAt > this.ttlMs) {
      this.removeJobDir(jobId);
      return null;
    }

    const fileEntry = job.files && job.files[format];
    if (!fileEntry || !fs.existsSync(fileEntry.filePath)) {
      return null;
    }

    return fileEntry;
  }

  cleanupExpiredJobs() {
    const now = Date.now();
    for (const [jobId, job] of this.jobs.entries()) {
      if (now - job.createdAt > this.ttlMs) {
        this.removeJobDir(jobId);
      }
    }
  }

  pruneOrphanDirs() {
    if (!fs.existsSync(this.tmpDir)) {
      return;
    }
    const now = Date.now();
    for (const entry of fs.readdirSync(this.tmpDir, { withFileTypes: true })) {
      if (!entry.isDirectory()) {
        continue;
      }
      const dirPath = path.join(this.tmpDir, entry.name);
      try {
        const stat = fs.statSync(dirPath);
        if (!this.jobs.has(entry.name) && now - stat.mtimeMs > this.ttlMs) {
          fs.rmSync(dirPath, { recursive: true, force: true });
        }
      } catch {
        // Ignore entries removed concurrently.
      }
    }
  }

  startCleanupScheduler(intervalMs = 5 * 60 * 1000) {
    if (this.cleanupTimer) {
      return;
    }
    this.cleanupTimer = setInterval(() => {
      this.cleanupExpiredJobs();
    }, intervalMs);
    this.cleanupTimer.unref();
  }

  stopCleanupScheduler() {
    if (this.cleanupTimer) {
      clearInterval(this.cleanupTimer);
      this.cleanupTimer = null;
    }
  }
}

module.exports = {
  JobStore,
  UUID_REGEX,
};

