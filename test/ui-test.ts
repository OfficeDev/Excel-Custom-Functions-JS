import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import { parseNumber } from "office-addin-cli";
import { AppType, startDebugging, stopDebugging } from "office-addin-debugging";
import { toOfficeApp } from "office-addin-manifest";
import { pingTestServer } from "office-addin-test-helpers";
import * as officeAddinTestServer from "office-addin-test-server";
import * as path from "path";
const host: string = "excel";
const manifestPath = path.resolve(`${process.cwd()}/test/test-manifest.xml`);
const port: number = 4201;
const testDataFile: string = `${process.cwd()}/test/src/test-data.json`;
const testJsonData = Object.assign({
        functions: {
            CLOCK: {
                result: {
                    amString: undefined,
                    pmString: undefined
                } 
            },
            INCREMENT: {
                result: undefined
            },
            LOG: {
                result: undefined
            }
        }
    }, JSON.parse(fs.readFileSync(testDataFile).toString()));
const testServer = new officeAddinTestServer.TestServer(port);
// let testValues: Array<{Value: any, Name: any}> & JSON = [];
let testValues: any = [];

describe("Test Excel Custom Functions", function () {
    before(`Setup test environment and sideload ${host}`, async function () {
        this.timeout(0);
        // Start test server and ping to ensure it's started
        const testServerStarted = await testServer.startTestServer(true /* mochaTest */);
        const serverResponse: { status: any, platform: any } = Object.assign({status: undefined, platform: undefined}, await pingTestServer(port));
        assert.strictEqual(testServerStarted, true);
        assert.strictEqual(serverResponse["status"], 200);

        // Call startDebugging to start dev-server and sideload
        const devServerCmd = `npm run dev-server -- --config ./test/webpack.config.js`;
        const devServerPort = parseNumber(process.env.npm_package_config_dev_server_port || 3000);
        await startDebugging(manifestPath, AppType.Desktop, toOfficeApp(host), undefined, undefined, 
            devServerCmd, devServerPort, undefined, undefined, undefined, false /* enableDebugging */);
    }),
    describe("Get test results for custom functions and validate results", function () {
        it("should get results from the taskpane application", async function () {
            this.timeout(0);
            // Expecting six result values
            testValues = Object.assign([], await testServer.getTestResults());
            assert.strictEqual(testValues.length, "6");
        });
        it("ADD function should return expected value", async function () {
            assert.strictEqual(testJsonData.functions.ADD.result, testValues[0].Value);
        });
        it("CLOCK function should return expected value", async function () {
            // Check that captured values are different to ensure the function is streaming
            assert.notStrictEqual(testValues[1].Value, testValues[2].Value);
            // Check if the returned string contains 'AM' or 'PM', indicating it's a time-stamp
            assert.strictEqual(true, testValues[1].Value.includes(testJsonData.functions.CLOCK.result.amString) || testValues[1].Value.includes(testJsonData.functions.CLOCK.result.pmString) ? true : false);
            assert.strictEqual(true, testValues[2].Value.includes(testJsonData.functions.CLOCK.result.amString) || testValues[2].Value.includes(testJsonData.functions.CLOCK.result.pmString) ? true : false);
        });
        it("INCREMENT function should return expected value", async function () {
            // Check that captured values are different to ensure the function is streaming
            assert.notStrictEqual(testValues[3].Value, testValues[4].Value);
            // Check to see that both captured streaming values are divisible by 4
            assert.strictEqual(0, testValues[3].Value % testJsonData.functions.INCREMENT.result);
            assert.strictEqual(0, testValues[4].Value % testJsonData.functions.INCREMENT.result);
        });
        it("LOG function should return expected value", async function () {
            assert.strictEqual(testJsonData.functions.LOG.result, testValues[5].Value);
        });
    });
    describe(`Get test results for Excel taskpane project and validate results`, function () {
        it("Validate expected result name", async function () {
            assert.strictEqual(testValues[6].Name, "fill-color");
        });
        it("Validate expected result", async function () {
            assert.strictEqual(testValues[6].Value, "#FFFF00");
        });
    });
    after("Teardown test environment", async function () {
        this.timeout(0);
        // Stop the test server
        const stopTestServer = await testServer.stopTestServer();
        assert.strictEqual(stopTestServer, true);

        // Unregister the add-in
        await stopDebugging(manifestPath);
    });
});