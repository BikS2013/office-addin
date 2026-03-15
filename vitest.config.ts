import { defineConfig } from "vitest/config";
import path from "path";

export default defineConfig({
  test: {
    globals: true,
    environment: "node",
    include: ["test_scripts/**/*.test.ts"],
    testTimeout: 10000,
  },
  resolve: {
    alias: {
      "@types": path.resolve(__dirname, "src/taskpane/types"),
      "@services": path.resolve(__dirname, "src/taskpane/services"),
    },
  },
});
