const fs = require('fs').promises;
const path = require('path');
const readlineSync = require('readline-sync');
const { OpenAI } = require('openai');
const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');

// Initialize OpenAI client
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY || readlineSync.question('Enter your OpenAI API key: ', { hideEchoBack: true })
});

async function readTxtFile(filePath) {
  try {
    return await fs.readFile(filePath, 'utf-8');
  } catch (error) {
    console.error(`Error reading TXT file: ${error.message}`);
    return null;
  }
}

async function readDocxFile(filePath) {
  try {
    const { value } = await mammoth.extractRawText({ path: filePath });
    return value;
  } catch (error) {
    console.error(`Error reading DOCX file: ${error.message}`);
    return null;
  }
}

async function readPdfFile(filePath) {
  try {
    const dataBuffer = await fs.readFile(filePath);
    const data = await pdfParse(dataBuffer);
    return data.text;
  } catch (error) {
    console.error(`Error reading PDF file: ${error.message}`);
    return null;
  }
}

async function readDocument(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  switch (ext) {
    case '.txt':
      return await readTxtFile(filePath);
    case '.docx':
      return await readDocxFile(filePath);
    case '.pdf':
      return await readPdfFile(filePath);
    default:
      console.error('Unsupported file format. Use TXT, DOCX, or PDF.');
      return null;
  }
}

async function generateTestCases(documentContent) {
  const prompt = `
You are a QA expert. Based on the following requirements document, generate detailed functional test cases in a structured format. Each test case should include:
- Test ID: Unique identifier (e.g., TC001)
- Category: Functional
- Description: Brief overview of the test case
- Test Steps: Step-by-step instructions
- Expected Result: Expected outcome
- Test Data: Specific inputs to be used

Requirements Document:
${documentContent}

Provide test cases in a clear, tabular plain text format.
  `;
  try {
    const response = await openai.chat.completions.create({
      model: 'gpt-4o',
      messages: [
        { role: 'system', content: 'You are a QA expert specializing in creating functional test cases.' },
        { role: 'user', content: prompt }
      ],
      max_tokens: 4000
    });
    return response.choices[0].message.content;
  } catch (error) {
    console.error(`Error generating test cases: ${error.message}`);
    return null;
  }
}

async function saveTestCases(testCases, outputDir = 'output') {
  try {
    await fs.mkdir(outputDir, { recursive: true });
    const timestamp = new Date().toISOString().replace(/[:.]/g, '');
    const outputFile = path.join(outputDir, `test_cases_${timestamp}.txt`);
    await fs.writeFile(outputFile, testCases, 'utf-8');
    console.log(`Test cases saved to ${outputFile}`);
    return outputFile;
  } catch (error) {
    console.error(`Error saving test cases: ${error.message}`);
    return null;
  }
}

async function main() {
  const filePath = readlineSync.question('Enter the path to the requirements document (TXT, DOCX, or PDF): ').trim();
  
  if (!await fs.access(filePath).then(() => true).catch(() => false)) {
    console.error('File not found. Please check the path and try again.');
    return;
  }

  const documentContent = await readDocument(filePath);
  if (!documentContent) {
    console.error('Failed to read document content.');
    return;
  }

  const testCases = await generateTestCases(documentContent);
  if (!testCases) {
    console.error('Failed to generate test cases.');
    return;
  }

  const outputFile = await saveTestCases(testCases);
  if (outputFile) {
    console.log(`Test cases successfully generated and saved to ${outputFile}`);
  } else {
    console.error('Failed to save test cases.');
  }
}

main().catch(error => console.error(`Error in main: ${error.message}`));