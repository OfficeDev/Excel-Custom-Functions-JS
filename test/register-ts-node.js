const path = require("path");
const tsNode = require("ts-node");

tsNode.register({
  files: true,
  project: path.resolve(__dirname, "tsconfig.json"),
});
