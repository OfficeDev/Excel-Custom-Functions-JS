module.exports = {
  root: true,
  parser: '@typescript-eslint/parser',
  plugins: [
    '@typescript-eslint',
    "excel-custom-functions"
  ],
  rules: {
      "excel-custom-functions/no-office-read-calls": "warn",
      "excel-custom-functions/no-office-write-calls": "error"
  },
  parserOptions: {
    project: "./tsconfig.json"
  }
};