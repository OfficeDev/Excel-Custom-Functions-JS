import functionsJsonData from "./test-data.json";
import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { addErrorResult, closeWorkbook, formatError, sleep } from "./test-helpers";

/* global Office, document, Excel, run, navigator */
const customFunctionsData = (<any>functionsJsonData).functions;
const port: number = 4201;
let testValues: any[] = [];

Office.onReady(async () => {
  document.getElementById("sideload-msg")!.style.display = "none";
  document.getElementById("app-body")!.style.display = "flex";
  document.getElementById("run")!.onclick = run;
  addTestResult("UserAgent", navigator.userAgent);

  try {
    const testServerResponse = (await pingTestServer(port)) as { status?: number };
    if (testServerResponse.status === 200) {
      await runCfTests();
      await sendTestResults(testValues, port);
      await closeWorkbook();
    } else {
      addErrorResult(testValues, `Ping failed: ${JSON.stringify(testServerResponse)}`);
      await sendTestResults(testValues, port).catch(() => {});
    }
  } catch (err) {
    testValues = [];
    addErrorResult(testValues, `Initialization failed: ${formatError(err)}`);
    await sendTestResults(testValues, port).catch(() => {});
  }
});

async function runCfTests(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      for (let key in customFunctionsData) {
        const formula: string = customFunctionsData[key].formula;
        const range = context.workbook.getSelectedRange();
        range.formulas = [[formula]];
        await context.sync();

        await sleep(5000);

        await readCFData(key, customFunctionsData[key].streaming !== undefined ? 2 : 1);
      }
    });
  } catch (err) {
    testValues = [];
    addErrorResult(testValues, `runCfTests failed: ${formatError(err)}`);
    throw err;
  }
}

export async function readCFData(cfName: string, readCount: number): Promise<void> {
  await Excel.run(async (context) => {
    for (let i = 0; i < readCount; i++) {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      await sleep(5000);

      addTestResult(cfName, range.values[0][0]);
    }
  });
}

function addTestResult(name: string, value: any) {
  const data = {
    name,
    value,
  };
  testValues.push(data);
}
