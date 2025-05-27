const fs = require('fs').promises;
const path = require('path');
const readlineSync = require('readline-sync');
const { OpenAI } = require('openai');
const mammoth = require('mammoth');
const XLSX = require('xlsx');

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY || readlineSync.question('Enter your OpenAI API key: ', { hideEchoBack: true })
});

async function readDocxFile(filePath) {
  try {
    const { value } = await mammoth.extractRawText({ path: filePath });
    return value;
  } catch (error) {
    console.error(`Error reading DOCX file: ${error.message}`);
    return null;
  }
}

function extractUserStoryTitles(userStoriesText) {
  const titles = [];
  const lines = userStoriesText.split('\n').map(line => line.trim());
  let currentTitle = null;

  for (const line of lines) {
    if (line.match(/^User Story \d+: .+/)) {
      currentTitle = line.replace(/^User Story \d+: /, '').trim();
      const prefix = currentTitle
        .split(' ')
        .map(word => word.charAt(0).toUpperCase())
        .join('')
        .slice(0, 5); // e.g., "Add a Single Product to Cart" -> "ASPTC"
      titles.push({ fullTitle: currentTitle, prefix });
    }
  }
  return titles.length > 0 ? titles : [{ fullTitle: 'Default', prefix: 'DEF' }];
}

async function generateTestCases(userStoriesText, storyTitles) {
  const titleMap = storyTitles.map(t => `Title: "${t.fullTitle}", Prefix: TC-${t.prefix}-<number>`).join('\n');
  const prompt = `
You are a QA expert tasked with generating exhaustive functional test cases for an e-commerce "Add to Cart" feature based on the provided user stories. For **each acceptance criterion** in every user story, generate **at least 3–5 test cases** to cover all scenarios, including:
- Positive cases (happy path, e.g., adding a valid product).
- Negative cases (e.g., out-of-stock products, invalid quantities).
- Edge cases (e.g., maximum/minimum quantities, special characters in inputs).
- Boundary cases (e.g., cart size limits, session timeouts).
- Accessibility scenarios (e.g., adding to cart using keyboard navigation).
- Database scenarios
- Security test scenarios.

Each test case must include:
- Test ID: Unique identifier in the format TC-<prefix>-<number>, where <prefix> is derived from the user story title (see below) and <number> is a sequential number (e.g., TC-ASPTC-001).
- Category: Functional
- Description: Brief overview of the test case, referencing the specific acceptance criterion and user story.
- Test Steps: Detailed, numbered step-by-step instructions (6–8 steps per test case, unless the scenario is inherently simple). Steps must include:
  - Preconditions (e.g., user logged in, product available in stock).
  - Specific user actions (e.g., clicking buttons, selecting product variants like size or color).
  - System verifications (e.g., checking for UI updates, error messages, or cart icon changes).
  - Example: Instead of "Navigate to product page," use "From the homepage, click the 'Clothing' category in the top navigation bar, then select the 'T-Shirt' product."
- Expected Result: Expected outcome, tied to the acceptance criterion.
- Test Data: Specific inputs (e.g., Product: T-Shirt, Size: M, Quantity: 1, Username: testuser).

User Story Titles and Test ID Prefixes:
${titleMap}

User Stories:
${userStoriesText}

Provide test cases in a plain text table format using pipes (|) to separate columns, with exactly 6 columns per row, no extra pipes, and no missing columns. Example:
| Test ID | Category | Description | Test Steps | Expected Result | Test Data |
|---------|----------|-------------|------------|-----------------|-----------|
| TC-ASPTC-001 | Functional | Validate single product addition (Add a Single Product to Cart) | 1. Open the e-commerce website in a browser. 2. Log in with a valid user account (username: testuser, password: password123). 3. Navigate to the "Clothing" category via the top navigation bar. 4. Click on the product "T-Shirt" to open its details page. 5. Verify the "Add to Cart" button is visible and enabled. 6. Select size "M" from the dropdown. 7. Enter "1" in the quantity field. 8. Click the "Add to Cart" button. | The T-Shirt is added to the cart, a success message "Item added to cart" is displayed, and the cart icon updates to show 1 item. | Product: T-Shirt, Size: M, Quantity: 1, Username: testuser, Password: password123 |
| TC-ASPTC-002 | Functional | Validate adding out-of-stock product (Add a Single Product to Cart) | 1. Open the e-commerce website in a browser. 2. Log in with a valid user account (username: testuser, password: password123). 3. Navigate to the "Clothing" category via the top navigation bar. 4. Select an out-of-stock product (e.g., Jacket). 5. Verify the "Add to Cart" button is disabled or shows an out-of-stock message. 6. Attempt to click the "Add to Cart" button. 7. Verify no item is added to the cart. | An error message "Product is out of stock" is displayed, and the cart remains unchanged. | Product: Jacket, Size: L, Quantity: 1, Username: testuser, Password: password123 |

Rules:
- Generate **at least 3–5 test cases per acceptance criterion** to ensure comprehensive coverage.
- Use the correct Test ID prefix for each user story (e.g., TC-ASPTC- for "Add a Single Product to Cart").
- Ensure Test IDs are unique and sequential within each prefix (e.g., TC-ASPTC-001, TC-ASPTC-002).
- Test steps must be detailed, actionable, and include specific user actions and system checks. Avoid vague steps.
- Cover all possible scenarios implied by the acceptance criteria, including edge cases and accessibility.
- Keep the table well-formatted with exactly 6 columns, no missing or extra columns.
  `;

  const retries = 3;
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const response = await openai.chat.completions.create({
        model: process.env.OPENAI_MODEL || 'gpt-4o',
        messages: [
          { role: 'system', content: 'You are a QA expert specializing in creating comprehensive functional test cases.' },
          { role: 'user', content: prompt }
        ],
        max_tokens: 8000
      });
      const testCasesText = response.choices[0].message.content;
      console.log('Raw OpenAI Response:\n', testCasesText); // Log for debugging
      if (!testCasesText || !testCasesText.includes('| Test ID')) {
        console.warn(`Attempt ${attempt}: Invalid or empty table format in response.`);
        if (attempt === retries) throw new Error('Failed to get valid table format after retries.');
        continue;
      }
      return testCasesText;
    } catch (error) {
      console.error(`Attempt ${attempt} failed: ${error.message}`);
      if (attempt === retries) {
        console.error('All retry attempts failed.');
        return null;
      }
    }
  }
  return null;
}

function parseTestCases(testCasesText) {
  if (!testCasesText || typeof testCasesText !== 'string') {
    console.error('Invalid or empty test cases text provided to parseTestCases.');
    return { headers: ['Test ID', 'Category', 'Description', 'Test Steps', 'Expected Result', 'Test Data'], testCases: [] };
  }

  const lines = testCasesText.split('\n').map(line => line.trim()).filter(line => line);
  const testCases = [];
  const headers = ['Test ID', 'Category', 'Description', 'Test Steps', 'Expected Result', 'Test Data'];
  let tableStart = false;

  for (const line of lines) {
    if (line.startsWith('| Test ID') && line.includes('Category') && line.includes('Description')) {
      tableStart = true;
      continue;
    }
    if (tableStart && line.startsWith('|') && !line.startsWith('|-')) {
      const columns = line.split('|').map(col => col.trim()).filter(col => col);
      if (columns.length !== headers.length) {
        console.warn(`Skipping malformed row: ${line}`);
        continue;
      }
      testCases.push(Object.fromEntries(headers.map((header, i) => [header, columns[i]])));
    }
  }
  console.log(`Parsed ${testCases.length} test cases.`);
  return { headers, testCases };
}

async function saveTestCasesToExcel(testCasesData, outputPath) {
  try {
    const { headers, testCases } = testCasesData;
    const worksheetData = [
      headers,
      ...testCases.map(tc => headers.map(header => tc[header]))
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Test Cases');

    worksheet['!cols'] = [
      { wch: 15 },  // Test ID
      { wch: 12 },  // Category
      { wch: 40 },  // Description
      { wch: 80 },  // Test Steps (increased for detailed steps)
      { wch: 50 },  // Expected Result
      { wch: 30 }   // Test Data
    ];

    headers.forEach((_, i) => {
      const cell = XLSX.utils.encode_cell({ r: 0, c: i });
      worksheet[cell].s = { font: { bold: true } };
    });

    await fs.writeFile(outputPath, XLSX.write(workbook, { type: 'buffer' }));
    console.log(`Test cases saved to ${outputPath}`);
    return outputPath;
  } catch (error) {
    console.error(`Error saving Excel file: ${error.message}`);
    return null;
  }
}

async function main() {
  const defaultProjectFolder = process.env.PROJECT_FOLDER || path.join(process.cwd(), 'project');
  console.log(`Default project folder: ${defaultProjectFolder}`);
  const projectFolder = readlineSync.question('Enter the project folder path to save the test cases document (press Enter for default): ').trim() || defaultProjectFolder;

  await fs.mkdir(projectFolder, { recursive: true });

  console.log('Reading user stories...');
  const userStoriesPath = readlineSync.question('Enter the path to the .docx file containing user stories: ').trim();
  if (!userStoriesPath.toLowerCase().endsWith('.docx') || !await fs.access(userStoriesPath).then(() => true).catch(() => false)) {
    console.error('Invalid or inaccessible .docx file. Please provide a valid path.');
    return;
  }

  const userStoriesText = await readDocxFile(userStoriesPath);
  if (!userStoriesText) {
    console.error('Failed to read user stories from the provided file.');
    return;
  }

  console.log('Generating test cases...');
  const storyTitles = extractUserStoryTitles(userStoriesText);
  const testCasesText = await generateTestCases(userStoriesText, storyTitles);
  if (!testCasesText) {
    console.error('Failed to generate test cases after retries.');
    return;
  }

  console.log('Parsing test cases...');
  const testCasesData = parseTestCases(testCasesText);
  if (!testCasesData || !testCasesData.testCases) {
    console.error('Failed to parse test cases: Invalid data structure returned.');
    return;
  }
  if (testCasesData.testCases.length === 0) {
    console.error('No valid test cases parsed from the response. Check the raw OpenAI response above for issues.');
    return;
  }
  console.log(`Successfully parsed ${testCasesData.testCases.length} test cases.`);

  const outputPath = path.join(projectFolder, 'AddToCartTestCases.xlsx');
  if (await fs.access(outputPath).then(() => true).catch(() => false)) {
    const overwrite = readlineSync.question('Output file already exists. Overwrite? (y/n): ').toLowerCase();
    if (overwrite !== 'y') {
      console.log('Operation cancelled.');
      return;
    }
  }

  console.log('Saving test cases to Excel...');
  const savedFile = await saveTestCasesToExcel(testCasesData, outputPath);
  if (savedFile) {
    console.log(`Test cases successfully saved to ${savedFile}`);
  } else {
    console.error('Failed to save test cases to Excel.');
  }
}

main().catch(error => console.error(`Error in main: ${error.message}`));