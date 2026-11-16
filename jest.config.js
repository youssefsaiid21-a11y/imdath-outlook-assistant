module.exports = {
  testEnvironment: "node",
  testMatch: ["**/tests/**/*.test.js"],
  testPathIgnorePatterns: ["\\.integration\\.js$", "\\.spec\\.js$"],
  collectCoverageFrom: [
    "server/**/*.js",
    "utils.js",
    "auth.js",
    "recipient.js",
    "context.js",
    "compose.js",
    "crm.js",
    "followups.js",
    "imdath-init.js",
    "!server/index.js",
  ],
  coverageDirectory: "coverage",
  verbose: true,
};
