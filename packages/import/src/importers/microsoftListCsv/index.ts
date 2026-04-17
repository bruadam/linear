import inquirer from "inquirer";
import type { Importer } from "../../types.ts";
import { MicrosoftListCsvImporter } from "./MicrosoftListCsvImporter.ts";

const BASE_PATH = process.cwd();

export const microsoftListCsvImport = async (): Promise<Importer> => {
  const answers = await inquirer.prompt<MicrosoftListImportAnswers>(questions);
  return new MicrosoftListCsvImporter(answers.microsoftListFilePath);
};

interface MicrosoftListImportAnswers {
  microsoftListFilePath: string;
}

const questions = [
  {
    basePath: BASE_PATH,
    type: "filePath",
    name: "microsoftListFilePath",
    message: "Select your exported CSV file of Microsoft List issues",
  },
];
