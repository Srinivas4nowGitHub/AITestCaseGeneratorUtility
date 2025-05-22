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
You are a QA expert tasked with generating exhaustive functional test cases for an e-commerce "Add to Cart" feature based on the provided user stories. For **each acceptance criterion** in every user story, generate one or more test cases to cover all scenarios, including positive, negative, and edge cases. Each test case must include:
- Test ID: Unique identifier in the format TC-<prefix>-<number>, where <prefix> is derived from the user story title (see below) and <number> is a sequential number (e.g., TC-ASPTC-001).
- Category: Functional
- Description: Brief overview of the test case, referencing the specific acceptance criterion and user story
- Test Steps: Numbered step-by-step instructions (e.g., 1. Do this, 2. Do that)
- Expected Result: Expected outcome, tied to the acceptance criterion
- Test Data: Specific inputs (e.g., Product: T-Shirt, Quantity: 1)

User Story Titles and Test ID Prefixes:
${titleMap}

User Stories:
${userStoriesText}

Provide test cases in a plain text table format using pipes (|) to separate columns, with exactly 6 columns per row, no extra pipes, and no missing columns. Example:
| Test ID | Category | Description | Test Steps | Expected Result | Test Data |
|---------|----------|-------------|------------|-----------------|-----------|
| TC-ASPTC-001 | Functional | Validate single product addition (Add a Single Product to Cart) | 1. Navigate to product page. 2. Click 'Add to Cart'. | Product added, success message shown. | Product: T-Shirt, Quantity: 1 |

Rules:
- Generate test cases for **every acceptance criterion** in each user story.
- Use the correct Test ID prefix for each user story (e.g., TC-ASPTC- for "Add a Single Product to Cart").
- Include positive tests (happy path), negative tests (error conditions), and edge cases (e.g., max/min quantities, out-of-stock).
- Ensure Test IDs are unique and sequential within each prefix (e.g., TC-ASPTC-001, TC-ASPTC-002).
- Do not skip any scenarios; cover all possible cases implied by the acceptance criteria.
- Keep the table well-formatted with no missing or extra columns.
  `;

  try {
    const response = await openai.chat.completions.create({
      model: 'gpt-4o',
      messages: [
        { role: 'system', content: 'You are a QA expert specializing in creating comprehensive functional test cases.' },
        { role: 'user', content: prompt }
      ],
      max_tokens: 8000
    });
    const testCasesText = response.choices[0].message.content;
    console.log('Raw OpenAI Response:\n', testCasesText); // Log for debugging
    return testCasesText;
  } catch (error) {
    console.error(`Error generating test cases: ${error.message}`);
    return null;
  }
}

function parseTestCases(testCasesText) {
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
      if (columns.length === headers.length) {
        testCases.push({
          'Test ID': columns[0],
          'Category': columns[1],
          'Description': columns[2],
          'Test Steps': columns[3],
          'Expected Result': columns[4],
          'Test Data': columns[5]
        });
      }
    }
  }
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

    // Set column widths for readability
    worksheet['!cols'] = [
      { wch: 15 },  // Test ID (wider for prefixes)
      { wch: 12 },  // Category
      { wch: 40 },  // Description (wider for clarity)
      { wch: 60 },  // Test Steps (wider for detailed steps)
      { wch: 50 },  // Expected Result
      { wch: 30 }   // Test Data
    ];

    // Add bold headers
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
  const defaultProjectFolder = path.join(process.cwd(), 'project');
  console.log(`Default project folder: ${defaultProjectFolder}`);
  const projectFolder = readlineSync.question('Enter the project folder path to save the test cases document (press Enter for default): ').trim() || defaultProjectFolder;

  await fs.mkdir(projectFolder, { recursive: true });

  const userStoriesPath = readlineSync.question('Enter the path to the .docx file containing user stories: ').trim();
  if (!await fs.access(userStoriesPath).then(() => true).catch(() => false)) {
    console.error('User stories file not found. Please check the path and try again.');
    return;
  }

  const userStoriesText = await readDocxFile(userStoriesPath);
  if (!userStoriesText) {
    console.error('Failed to read user stories from the provided file.');
    return;
  }

  const storyTitles = extractUserStoryTitles(userStoriesText);
  const testCasesText = await generateTestCases(userStoriesText, storyTitles);
  if (!testCasesText) {
    console.error('Failed to generate test cases.');
    return;
  }

  const testCasesData = parseTestCases(testCasesText);
  if (testCasesData.testCases.length === 0) {
    console.error('No valid test cases parsed from the response. Check the raw OpenAI response above for issues.');
    return;
  }

  const outputPath = path.join(projectFolder, 'AddToCartTestCases.xlsx');
  const savedFile = await saveTestCasesToExcel(testCasesData, outputPath);
  if (savedFile) {
    console.log(`Test cases successfully saved to ${savedFile}`);
  } else {
    console.error('Failed to save test cases to Excel.');
  }
}

main().catch(error => console.error(`Error in main: ${error.message}`));