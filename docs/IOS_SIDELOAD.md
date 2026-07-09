# Installing on your iPhone with a free Apple ID (no paid account, no Mac)

This app can't use Expo Go (it has a custom native module, `react-native-mmkv`), and a free
Apple ID can't drive an EAS device build. The free path is:

1. **Build an unsigned `.ipa`** on GitHub's macOS runners (the `iOS unsigned IPA` workflow) —
   no Apple account, no secrets, no local Xcode, still on Expo SDK 57.
2. **Sideload it** onto your iPhone from your **Windows PC** with **Sideloadly** + your **free
   Apple ID** — Sideloadly signs it on the way in.

Nothing in the app changes — the workflow only generates the native iOS project (`expo
prebuild`) and compiles it.

---

## What you need to set up (one-time)

- **GitHub secrets/settings:** **none.** An unsigned build needs no certificates or secrets.
  Actions is enabled by default on your repo. If it was ever turned off: GitHub → your repo →
  **Settings → Actions → General → Allow all actions**, and under **Workflow permissions** the
  defaults are fine (this workflow doesn't push anything).
- **macOS runner minutes:** GitHub-hosted macOS runners are free for public repos; for a
  **private** repo they draw from your plan's included minutes (Free plan ≈ 2,000 min/month,
  and macOS counts 10×, so ≈ 200 macOS-minutes — enough for several builds). You only rebuild
  when native code/deps change, not weekly.

That's it — I can't click these for you, but there's nothing to configure for a first run.

---

## Step 1 — Trigger the build on GitHub

The workflow has **two** triggers; use whichever fits:

### Option A — Tag trigger (works right now, from any branch)

The GitHub "Run workflow" button only appears once the workflow file is on your **default
branch (`main`)**. Until then, trigger it by pushing a tag whose name starts with `ios-build-`:

```bash
git checkout claude/personal-fitness-tracker-72pwhh   # the branch with the app + workflow
git pull
git tag ios-build-1
git push origin ios-build-1
```

Every time you want a fresh build, push a new tag (`ios-build-2`, `ios-build-3`, …).

### Option B — Manual button (after the workflow is on `main`)

Once `.github/workflows/ios-ipa.yml` is merged to `main`:

1. GitHub → your repo → **Actions** tab.
2. Left sidebar → **iOS unsigned IPA**.
3. **Run workflow ▸** → pick the branch to build (e.g. `main` or the feature branch) → **Run
   workflow**.

Either way, the run appears under **Actions**. It takes roughly **15–30 minutes** (Skia +
Reanimated compile). A green check = success.

---

## Step 2 — Download the IPA artifact

1. GitHub → **Actions** → click the finished run.
2. Scroll to the bottom, **Artifacts** section → **`Fitness-unsigned-ipa`** → click to
   download. (Artifacts are kept 30 days.)
3. It downloads as **`Fitness-unsigned-ipa.zip`**. **Unzip it** — inside is
   **`Fitness-unsigned.ipa`**. That `.ipa` is the file you sideload.

> GitHub always wraps artifacts in a `.zip`. Don't sideload the `.zip` — extract the `.ipa`
> first.

---

## Step 3 — Install on your iPhone with Sideloadly (Windows + free Apple ID)

**One-time Windows setup**

1. Install **iTunes** and **iCloud** from **apple.com** (the direct downloads, *not* the
   Microsoft Store versions — Sideloadly needs Apple's device drivers).
2. Download **Sideloadly** from **https://sideloadly.io** and install it.
3. *(Recommended)* If your Apple ID has two-factor auth (it should), create an **app-specific
   password**: **appleid.apple.com → Sign-In & Security → App-Specific Passwords → +**. You'll
   paste this instead of your normal password when Sideloadly asks.

**Install the app**

4. Connect your iPhone to the PC with a USB cable. Unlock it and tap **Trust** on the
   "Trust This Computer?" prompt.
5. Open **Sideloadly**. It should show your iPhone at the top. If not, open iTunes once so it
   recognizes the device, then reopen Sideloadly.
6. Drag **`Fitness-unsigned.ipa`** into the Sideloadly window (or click the file box and pick
   it).
7. In **Apple account**, enter your **free Apple ID email**. Click **Start**.
8. When prompted for the password, enter your **app-specific password** (or your Apple ID
   password if you don't use 2FA).
9. Sideloadly signs and installs. First run can take a minute or two.

**Trust the app on the iPhone (one-time per certificate)**

10. On the iPhone: **Settings → General → VPN & Device Management** → under *Developer App*,
    tap your Apple ID email → **Trust** → confirm.
11. Launch **Fitness** from your home screen. Done.

---

## Living with a free Apple ID (the trade-offs)

- **7-day expiry.** A free-signed app stops launching after **7 days**. To refresh: reconnect
  the iPhone and **re-run Sideloadly** with the same `.ipa` and Apple ID — this re-signs in
  place.
  - **Your data is safe.** Re-signing/reinstalling the same app (same bundle id
    `com.anas.fitness`, same Apple ID) **keeps your workouts, meals, measurements, and
    photos** — the local SQLite database is untouched. Only *deleting* the app wipes data (so
    take a **Backup** from Settings → Backup before any big change, per the app's own data-safety
    design).
  - **Want to avoid the weekly re-run?** Use **AltStore** or **SideStore** instead of
    Sideloadly (same idea, same free Apple ID) — they **auto-refresh** the 7-day signature over
    Wi-Fi. Slightly more setup; worth it for daily use.
- **3-app limit** per free Apple ID, and you must be on the same network for auto-refresh.
- **You only rebuild the `.ipa`** (Steps 1–2) when you change native code or dependencies —
  not for the weekly refresh, and not for normal use.

---

## When you change native dependencies

If you add/upgrade a native module (anything with iOS pods), push a new `ios-build-N` tag (or
hit Run workflow), download the new artifact, and sideload it over the top. JS-only changes
don't need a new build unless you're distributing them — for live JS iteration you'd use a dev
build + `npx expo start`, which is the paid-account path.

---

## If a build ever fails

Open the failed run under **Actions** and read the red step:

- **`pod install` fails** — usually a transient CocoaPods CDN hiccup; re-run the job.
- **Archive fails on signing** — the workflow already disables signing
  (`CODE_SIGNING_ALLOWED=NO`); if Xcode still complains about a specific pod, it's almost always
  a runner-image change — re-run, and if it persists, tell me the failing step and I'll adjust.
- **Xcode version** — the workflow pins `macos-15` (Xcode 16) for SDK 57 / RN 0.81; if GitHub
  retires that image, bump the `runs-on:` line.
