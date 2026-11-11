/** Independent page preparation; only the commit callback may touch the live deck. */
export class PageFillQueue<T, R> {
  private active = new Map<number, { key: number; prepare: () => Promise<T>; resolve: (value: R) => void; reject: (error: unknown) => void; value?: T }>();
  private waiting: Array<{ key: number; prepare: () => Promise<T>; resolve: (value: R) => void; reject: (error: unknown) => void; value?: T }> = [];
  private preparing = 0;
  private committing = false;
  private closed = false;
  private peak = 0;
  constructor(private commit: (values: T[]) => Promise<R[]>) {}
  get pending(): boolean { return !!(this.active.size || this.waiting.length); }
  get peakPreparing(): number { return this.peak; }
  run(key: number, prepare: () => Promise<T>): Promise<R> {
    if (this.closed) return Promise.reject(new Error('PAGE_CLOSED: page queue disposed'));
    if (this.active.has(key) || this.waiting.some(job => job.key === key)) return Promise.reject(new Error(`PAGE_BUSY: page ${key} is already being prepared`));
    return new Promise<R>((resolve, reject) => {
      this.waiting.push({key, prepare, resolve, reject});
      this.pump();
    });
  }
  dispose(): void {
    this.closed = true;
    const error = new Error('PAGE_CLOSED: page queue disposed');
    for (const job of this.waiting.splice(0)) job.reject(error);
    for (const job of this.active.values()) job.reject(error);
  }
  private pump(): void {
    if (this.closed || this.committing) return;
    while (this.active.size < 3 && this.waiting.length) {
      const job = this.waiting.shift()!;
      this.active.set(job.key, job); this.preparing++;
      this.peak = Math.max(this.peak, this.preparing);
      void job.prepare().then(value => {job.value = value;}, error => {
        this.active.delete(job.key); job.reject(error);
        // A required page failed. Do not launch work already waiting behind it.
        for (const queued of this.waiting.splice(0)) queued.reject(error);
      }).finally(() => {this.preparing--; void this.flush();});
    }
  }
  private async flush(): Promise<void> {
    if (this.closed || this.committing || this.preparing) return;
    const batch = [...this.active.values()];
    if (!batch.length) {this.pump(); return;}
    this.committing = true;
    try {
      const results = await this.commit(batch.map(job => job.value!));
      if (results.length !== batch.length) throw new Error('Page commit receipt count mismatch');
      batch.forEach((job, index) => job.resolve(results[index]));
    } catch (error) {
      for (const job of batch) job.reject(error);
      for (const job of this.waiting.splice(0)) job.reject(error);
    } finally {
      for (const job of batch) this.active.delete(job.key);
      this.committing = false; this.pump();
    }
  }
}
