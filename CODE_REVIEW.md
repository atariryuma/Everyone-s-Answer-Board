# Code Review Notes

## Findings

1. **POST request parsing lacks basic input validation**
   - `doPost` blindly calls `JSON.parse` on the entire `e.postData.contents` without checking content type, bounding payload size, or handling malformed JSON separately. A large or invalid request body will throw before authentication runs and returns only a generic error, making the endpoint easy to DoS and hard to debug. Add explicit content-type checks, size limits, and a guarded parse path that returns a structured 400-style response on bad input.
2. **Client-controlled publish payload is saved without validation or concurrency safety**
   - The `publishApp` action spreads `request.config` straight into `publishConfig` and persists it via `saveUserConfig` with no schema validation or field allowlist. Clients can therefore mutate sensitive settings (e.g., `spreadsheetId`, `displaySettings`, or other stored flags) and overwrite concurrent admin changes because there is no etag or version check. Introduce strict validation/allowlisting and optimistic concurrency to prevent privilege or integrity issues.
3. **Security log persistence can silently exceed property quotas**
   - `persistSecurityLog` stores every high/critical event in Script Properties under a timestamped key, relying on a best-effort cleanup that only runs after writes. Under sustained error conditions this will still perform two property writes per request and can quickly exhaust the 500k character quota, causing later writes to fail silently and losing auditability. Consider batching logs, using a bounded queue before writes, and short-circuiting when storage nears quota.
