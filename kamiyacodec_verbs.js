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
// Function to write data to a JSON file
function writeData(filePath, data) {
  console.info(`Writing data to ${filePath}`);
  try {
      // Ensure the directory exists
      const dir = path.dirname(filePath);
      if (!fs.existsSync(dir)) {
          fs.mkdirSync(dir, { recursive: true });
      }

      // Serialize data to JSON and write to the file
      const jsonData = JSON.stringify(data, null, 2); // Pretty print JSON
      fs.writeFileSync(filePath, jsonData, 'utf8');
      console.debug('Data written successfully');
  } catch (error) {
      console.error(`Error writing data to file: ${error}`);
      throw error;
  }
}

function fetchConjugations(verb, type) {
  const forms = ["Negative", "Conjunctive", "Dictionary", "Conditional", "Imperative", "Volitional", "Te", "Ta", "Tara", "Tari", "Zu"];
  const auxiliaries = ["Potential", "Masu", "Nai", "Tai", "Tagaru", "Hoshii", "Rashii", "SoudaHearsay", "SoudaConjecture", "SeruSaseru", "ShortenedCausative", "ReruRareru", "CausativePassive", "ShortenedCausativePassive", "Ageru", "Sashiageru", "Yaru", "Morau", "Itadaku", "TeIru", "TeAru", "Miru", "Kureru", "Kudasaru", "Iku", "Kuru", "Oku", "Shimau"];

  const conjugations = {};

  forms.forEach(form => {
    try {
      const results = codec.conjugate(verb, form, type === 'TypeI'? false : true);
      conjugations[form] = results.join(', ');
    } catch (error) {
      console.error(`Error conjugating ${verb} (${form}): ${error.message}`);
    }
  });

  auxiliaries.forEach(aux => {
    forms.forEach(form => {
        try {
          const results = codec.conjugateAuxiliaries(verb, [aux], form, type === 'TypeI'? false : true);
          conjugations[`${aux} (${form})`] = results.join(', ');
        } catch (error) {
          console.error(`Error in auxiliary conjugation ${verb} + ${aux} (${form}): ${error.message}`);
        }
    }); // This was the missing closing bracket
  });

  return conjugations;
}


// Main function to handle the workflow
async function main() {
  try {
    const inputFile = path.join("C:\\Projects\\Kamiya\\input", 'input_verb.json');
    const outputFile = path.join("C:\\Projects\\Kamiya\\output", 'output_verb.json');

    const words = readData(inputFile);  // This file should contain an array of words
    const conjugatedWords = words.map(word => {
      const conjugations = fetchConjugations(word, 'TypeI');  // or 'TypeII' depending on the verb type
      return { word, conjugations };
    });
    writeData(outputFile, conjugatedWords);  // Save the conjugated words to output file

    console.info('Process completed successfully');
  } catch (error) {
    console.error(`An error occurred in the main function: ${error}`);
  }
}
main().catch(error => console.error(error));
