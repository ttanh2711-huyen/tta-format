/** @type {import('jest').Config} */
module.exports = {
  testEnvironment: "node",
  roots: ["<rootDir>/src", "<rootDir>/tools"],
  testMatch: ["**/*.test.ts", "**/*.test.js"],
  transform: {
    "^.+\\.ts$": ["ts-jest", { tsconfig: "tsconfig.jest.json" }],
  },
};
