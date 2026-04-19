// @ts-check
import { defineConfig } from "astro/config";

import react from "@astrojs/react";
import tailwindcss from "@tailwindcss/vite";

// https://astro.build/config
export default defineConfig({
  site: "https://danielth-uk.github.io",
  base: "/excel-exams-workbook-generator",
  integrations: [react()],

  vite: {
    plugins: [tailwindcss()],
  },
});
