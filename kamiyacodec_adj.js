const fs = require('fs');
const path = require('path');
const codec = require('kamiya-codec');

function readData(filePath) {
    console.info(`Reading data from ${filePath}`);
    try {
        const jsonData = fs.readFileSync(filePath, 'utf8');
        return JSON.parse(jsonData);
    } catch (error) {
        console.error(`Error reading data from file: ${error}`);
        throw error;
    }
}

function writeData(filePath, data) {
  console.info(`Writing data to ${filePath}`);
  try {
      const dir = path.dirname(filePath);
      if (!fs.existsSync(dir)) {
          fs.mkdirSync(dir, { recursive: true });
      }

      const jsonData = JSON.stringify(data, null, 2);
      fs.writeFileSync(filePath, jsonData, 'utf8');
      console.debug('Data written successfully');
  } catch (error) {
      console.error(`Error writing data to file: ${error}`);
      throw error;
  }
}

function fetchConjugations(adjective) {
  const forms = ["Negative", "Conditional", "Tari", "Present", "Prenomial", "Past", "NegativePast", "ConjunctiveTe", "Adverbial", "TaraConditional", "Noun", "StemSou", "StemNegativeSou"];
  
  const conjugations = {};
  
  forms.forEach(form => {
    try {
      console.log(`Conjugating ${adjective} as I-adjective (${form})`);
      const results = codec.adjConjugate(adjective, form, true);  // Assuming all are I-adjectives
      conjugations[form] = results ? results.join(', ') : 'No result';
    } catch (error) {
      console.error(`Error conjugating ${adjective} (${form}): ${error.message}`);
      conjugations[form] = 'Error';  // Indicate an error in output for troubleshooting
    }
  });

  return conjugations;
}

async function main() {
  try {
    const inputFile = path.join("C:\\Projects\\Kamiya\\input", 'input_adj.json');
    const outputFile = path.join("C:\\Projects\\Kamiya\\output", 'output_adj.json');

    const adjectives = readData(inputFile); // Expected to be an array of strings
    const conjugatedAdjectives = adjectives.map(adjective => {
      const conjugations = fetchConjugations(adjective);
      return { word: adjective, conjugations };
    });
    writeData(outputFile, conjugatedAdjectives);

    console.info('Process completed successfully');
  } catch (error) {
    console.error(`An error occurred in the main function: ${error}`);
  }
}
main().catch(error => console.error(error));
