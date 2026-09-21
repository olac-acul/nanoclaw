import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['.claude/skills/add-outlook-readonly/outlook-readonly.test.ts'],
  },
});
