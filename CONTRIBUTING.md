# VMS PCF — Developer Guide (monorepo + shared solution)

This repository is a **monorepo** that holds all of the VMS PCF (Power Apps Component
Framework) components together with **one shared Dataverse solution** (`VMS_PCF`) that packages
all of them into a single deployable.

> **TL;DR**
> - One solution → many controls.
> - Each component is its **own folder / project** (its own code, manifest, build).
> - Everyone works in the **same repo** with branches/PRs. Nobody imports their own component
>   separately — one build of the solution bundles them all.

---

## 1. Repository layout

```
PCF/                                  (repo root, git)
├── DataExport/                       (component #1 — PCF project)
│   ├── DataExport/                   control: ControlManifest.Input.xml, index.ts, css/
│   ├── DataExport.pcfproj
│   ├── package.json / pcfconfig.json / tsconfig.json
│   └── .gitignore                    (node_modules, out, bin, obj)
│
├── VesselScheduleGrid/               (each new component = a new top-level folder)
│   └── ...
│
└── Solution/                         (the shared Dataverse solution)
    ├── Solution.cdsproj              (one <ProjectReference> per component)
    └── src/Other/Solution.xml        (UniqueName = VMS_PCF, publisher, version)
```

### Key concepts

- **One solution = many controls.** Components are **not** written in a single shared
  codebase. Each one stays self-contained in its own folder and the solution collects them
  via `<ProjectReference>` entries.
- **Shared across all components:**
  - Namespace `VMS` → controls become `VMS.DataExport`, `VMS.VesselScheduleGrid`,
    etc. (grouped under one product).
  - Publisher `VMSControls` / prefix `vms`.
  - One solution: **`VMS_PCF`**.
- **Versioning (important — read before every deploy):**
  - Each control has its **own `version`** in its `ControlManifest.Input.xml` — bump it when
    that specific component changes.
  - The **solution** has its own `Version` in `Solution/src/Other/Solution.xml`. **Bump it
    before every deployment.** If you don't raise the version, the pipeline treats it as the
    same build and will not upgrade the target environment (it does a conflicting/no-op import
    instead of an upgrade). Use a simple scheme, e.g. `1.0.0.0 → 1.0.0.1`.

---

## 2. How to add a new component to this solution

This is the main workflow. Run everything from the **repo root**.

### Step 1 — Scaffold the new PCF project

```bash
pac pcf init --namespace VMS --name VesselScheduleGrid --template field \
  --outputDirectory VesselScheduleGrid --run-npm-install
```

- Keep `--namespace VMS` so all controls stay grouped under the same product.
- `--template field` → control bound to a single field; `--template dataset` → grid-style
  control bound to a set of records.
- Add `--framework react` for a React + Fluent UI control (default is plain HTML).
- `--run-npm-install` installs dependencies automatically.

### Step 2 — Apply the TypeScript build fix (required for every new component)

In the new component's `tsconfig.json`, inside `compilerOptions`, set:

```jsonc
"module": "esnext",
"moduleResolution": "bundler"
```

Without this, the production / solution build fails with `TS5103`. The cause is the
deprecated `moduleResolution: "node"` inherited from the pcf-scripts base tsconfig; the
`bundler` resolution is the clean fix. (The same fix is already applied in
`DataExport`.)

### Step 3 — Register the component in the shared solution

In `Solution/Solution.cdsproj`, add one line to the same `<ItemGroup>` that already holds the
existing reference:

```xml
<ItemGroup>
  <ProjectReference Include="..\DataExport\DataExport.pcfproj" />
  <ProjectReference Include="..\VesselScheduleGrid\VesselScheduleGrid.pcfproj" />   <!-- NEW -->
</ItemGroup>
```

### Step 4 — Build the whole solution

```bash
# From the repo root. Bundles ALL referenced controls into ONE solution zip.
dotnet build Solution/Solution.cdsproj -c Debug     # → unmanaged (for DEV environments)
dotnet build Solution/Solution.cdsproj -c Release   # → managed (for test / prod)
```

The output zip is written to `Solution/bin/<Debug|Release>/`.

That's it — the new control is now part of the `VMS_PCF` solution and ships with the next
import.

---

## 3. Local development of a single component

```bash
git clone <repo>
cd <ComponentFolder>
npm install
npm start watch        # local test harness with mock data
```

> Note: `context.webAPI` / real Dataverse queries do **not** work in the local harness — they
> are only exercised after deployment.

---

## 4. Team workflow (monorepo)

- **One git repo.** Each component has its own `.gitignore` (from `pac pcf init`) that excludes
  `node_modules/`, `out/`, `bin/`, `obj/`. Built solution zips should not be committed — add
  them to a root `.gitignore` (e.g. `*.zip`).
- **Branch per feature/component → PR → merge.**
  - Different components live in different folders, so two developers working on different
    components **never conflict**.
  - Two people on the same component is a normal git merge.
- **Nobody imports their component separately.** One build of the solution produces one zip
  with all controls, and that single zip is deployed (see below).

---

## 5. Deployment / ALM (Power Platform Pipelines)

We promote the solution **DEV → TEST → PROD** using in-product **Power Platform Pipelines**.

**Golden rule:** the **unmanaged** solution lives only in **DEV**; **TEST and PROD only ever
receive the managed** version, produced by the pipeline. Never import unmanaged into TEST/PROD.

1. **Build from the repo** (the source of truth — PCF code is *not* edited inside Dataverse):
   ```bash
   dotnet build Solution/Solution.cdsproj -c Debug     # → unmanaged zip (for DEV)
   ```
   Output: `Solution/bin/Debug/*.zip`.
2. **Import the unmanaged zip into DEV** (make.powerapps.com → Solutions → Import) and test.
3. **Run the pipeline** from DEV: it converts the solution to **managed** and deploys it to
   TEST, then PROD. No manual managed build needed for downstream — the pipeline does it.

**Before each deploy, bump the solution `Version`** in `Solution/src/Other/Solution.xml`
(see §1). Same `UniqueName` (`VMS_PCF`) and publisher prefix (`vms`) must stay identical across
all environments, otherwise the pipeline creates a duplicate solution instead of upgrading.

> Power Platform Pipelines requires the TEST/PROD environments to be **Managed Environments**
> (premium licensing).

---
