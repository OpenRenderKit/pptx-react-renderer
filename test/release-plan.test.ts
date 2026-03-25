/* @vitest-environment node */

import { describe, expect, it } from "vitest";

import { getReleasePlan } from "../scripts/release-plan.mjs";

describe("getReleasePlan", () => {
  it("keeps the source version for the first publish", () => {
    expect(getReleasePlan("0.1.0", "release:minor", true)).toMatchObject({
      nextVersion: "0.1.0",
      effectiveBump: "initial",
      shouldPublish: true,
      isInitialRelease: true,
    });
  });

  it("maps release:major to a minor bump while pre-1.0", () => {
    expect(getReleasePlan("0.2.3", "release:major")).toMatchObject({
      nextVersion: "0.3.0",
      effectiveBump: "minor",
    });
  });

  it("maps release:minor to a patch bump while pre-1.0", () => {
    expect(getReleasePlan("0.2.3", "release:minor")).toMatchObject({
      nextVersion: "0.2.4",
      effectiveBump: "patch",
    });
  });

  it("uses standard semver once 1.0.0 has been reached", () => {
    expect(getReleasePlan("1.2.3", "release:major")).toMatchObject({
      nextVersion: "2.0.0",
      effectiveBump: "major",
    });

    expect(getReleasePlan("1.2.3", "release:minor")).toMatchObject({
      nextVersion: "1.3.0",
      effectiveBump: "minor",
    });
  });

  it("supports skipping a publish", () => {
    expect(getReleasePlan("0.2.3", "release:skip")).toMatchObject({
      nextVersion: "0.2.3",
      effectiveBump: "skip",
      shouldPublish: false,
    });
  });
});

