const ALLOWED_LABELS = [
  "release:patch",
  "release:minor",
  "release:major",
  "release:skip",
];

export function parseVersion(version) {
  const match = /^(\d+)\.(\d+)\.(\d+)$/.exec(version);

  if (!match) {
    throw new Error(`Invalid semver version: ${version}`);
  }

  return {
    major: Number(match[1]),
    minor: Number(match[2]),
    patch: Number(match[3]),
  };
}

export function bumpVersion(version, bump) {
  const parsed = parseVersion(version);

  switch (bump) {
    case "major":
      return `${parsed.major + 1}.0.0`;
    case "minor":
      return `${parsed.major}.${parsed.minor + 1}.0`;
    case "patch":
      return `${parsed.major}.${parsed.minor}.${parsed.patch + 1}`;
    default:
      throw new Error(`Unsupported bump type: ${bump}`);
  }
}

export function getReleasePlan(currentVersion, requestedLabel, isInitialRelease = false) {
  if (!ALLOWED_LABELS.includes(requestedLabel)) {
    throw new Error(
      `Unsupported release label: ${requestedLabel}. Expected one of ${ALLOWED_LABELS.join(", ")}.`,
    );
  }

  if (requestedLabel === "release:skip") {
    return {
      currentVersion,
      nextVersion: currentVersion,
      requestedLabel,
      effectiveBump: "skip",
      shouldPublish: false,
      isInitialRelease,
    };
  }

  if (isInitialRelease) {
    return {
      currentVersion,
      nextVersion: currentVersion,
      requestedLabel,
      effectiveBump: "initial",
      shouldPublish: true,
      isInitialRelease,
    };
  }

  const { major } = parseVersion(currentVersion);

  let effectiveBump;
  if (requestedLabel === "release:major") {
    effectiveBump = major === 0 ? "minor" : "major";
  } else if (requestedLabel === "release:minor") {
    effectiveBump = major === 0 ? "patch" : "minor";
  } else {
    effectiveBump = "patch";
  }

  return {
    currentVersion,
    nextVersion: bumpVersion(currentVersion, effectiveBump),
    requestedLabel,
    effectiveBump,
    shouldPublish: true,
    isInitialRelease,
  };
}

function printEnv(plan) {
  console.log(`CURRENT_VERSION=${plan.currentVersion}`);
  console.log(`NEXT_VERSION=${plan.nextVersion}`);
  console.log(`REQUESTED_LABEL=${plan.requestedLabel}`);
  console.log(`EFFECTIVE_BUMP=${plan.effectiveBump}`);
  console.log(`SHOULD_PUBLISH=${plan.shouldPublish}`);
  console.log(`IS_INITIAL_RELEASE=${plan.isInitialRelease}`);
}

function main() {
  const args = process.argv.slice(2);
  const currentVersion = args[0];
  const requestedLabel = args[1];
  const format = args.includes("--format=env") ? "env" : "json";
  const isInitialRelease = args.includes("--initial");

  if (!currentVersion || !requestedLabel) {
    console.error(
      "Usage: node scripts/release-plan.mjs <current-version> <release:patch|release:minor|release:major|release:skip> [--initial] [--format=env]",
    );
    process.exit(1);
  }

  const plan = getReleasePlan(currentVersion, requestedLabel, isInitialRelease);

  if (format === "env") {
    printEnv(plan);
    return;
  }

  console.log(JSON.stringify(plan, null, 2));
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main();
}

