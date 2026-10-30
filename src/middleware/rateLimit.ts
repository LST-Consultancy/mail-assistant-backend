import type { RequestHandler } from "express";

// shortcut: in-memory counters, per process. Move to Redis when running >1 replica.
export class UserLimits {
  private minute = new Map<string, { window: number; count: number }>();
  private daily = new Map<string, { day: string; tokens: number }>();

  constructor(
    private perMinute: number,
    private dailyBudget: number,
    private now: () => number = Date.now,
  ) {}

  /** Returns false when the user exceeded the per-minute request limit. */
  hit(user: string): boolean {
    const window = Math.floor(this.now() / 60_000);
    const e = this.minute.get(user);
    if (!e || e.window !== window) {
      this.minute.set(user, { window, count: 1 });
      return true;
    }
    return ++e.count <= this.perMinute;
  }

  private today(user: string) {
    const day = new Date(this.now()).toISOString().slice(0, 10);
    let e = this.daily.get(user);
    if (!e || e.day !== day) this.daily.set(user, (e = { day, tokens: 0 }));
    return e;
  }

  hasBudget(user: string): boolean {
    return this.today(user).tokens < this.dailyBudget;
  }

  addTokens(user: string, tokens: number): void {
    this.today(user).tokens += tokens;
  }
}

export function rateLimit(limits: UserLimits): RequestHandler {
  return (req, res, next) => {
    const user = req.user!.idHash;
    if (!limits.hit(user)) {
      res.setHeader("Retry-After", "60");
      return void res.status(429).json({ error: "Too many requests. Please wait a minute." });
    }
    if (!limits.hasBudget(user)) {
      return void res.status(429).json({ error: "Daily AI usage limit reached. Try again tomorrow." });
    }
    next();
  };
}
