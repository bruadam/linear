/* eslint-disable no-console */
import { LinearClient } from "@linear/sdk";
import csv from "csvtojson";
import inquirer from "inquirer";
import ora from "ora";
import type { IssuePriority, IssueStatus, Importer } from "../../types.ts";
import { MicrosoftListCsvImporter, type ColumnMapping, type DateFormat } from "./MicrosoftListCsvImporter.ts";

const BASE_PATH = process.cwd();

const SKIP = "__skip__";
const KEEP = "__keep__";

const STATUS_TYPE_CHOICES: { name: string; value: IssueStatus }[] = [
  { name: "Backlog", value: "backlog" },
  { name: "Unstarted (Todo)", value: "unstarted" },
  { name: "Started (In Progress)", value: "started" },
  { name: "Completed (Done)", value: "completed" },
  { name: "Canceled", value: "canceled" },
];

const PRIORITY_CHOICES: { name: string; value: IssuePriority }[] = [
  { name: "No priority", value: 0 },
  { name: "Urgent", value: 1 },
  { name: "High", value: 2 },
  { name: "Medium", value: 3 },
  { name: "Low", value: 4 },
];

const DATE_FORMAT_CHOICES: { name: string; value: DateFormat }[] = [
  { name: "Auto-detect", value: "auto" },
  { name: "DD-MM-YYYY  (e.g. 24-07-2026)", value: "DD-MM-YYYY" },
  { name: "MM-DD-YYYY  (e.g. 07-24-2026)", value: "MM-DD-YYYY" },
  { name: "YYYY-MM-DD  (e.g. 2026-07-24)", value: "YYYY-MM-DD" },
  { name: "DD/MM/YYYY  (e.g. 24/07/2026)", value: "DD/MM/YYYY" },
  { name: "MM/DD/YYYY  (e.g. 07/24/2026)", value: "MM/DD/YYYY" },
];

export const microsoftListCsvImport = async (apiKey: string, apiUrl?: string): Promise<Importer> => {
  // ── Step 1: select the CSV file ─────────────────────────────────────────
  const { filePath } = await inquirer.prompt([
    {
      basePath: BASE_PATH,
      type: "filePath",
      name: "filePath",
      message: "Select your exported CSV file of Microsoft List issues",
    },
  ]);

  // ── Step 2: parse CSV to discover columns and a sample row ───────────────
  const rows = (await csv().fromFile(filePath)) as Record<string, string>[];
  if (rows.length === 0) {
    throw new Error("The selected CSV file contains no rows.");
  }

  const columns = Object.keys(rows[0]);
  const sampleRow = rows[0];

  console.log(`\nDetected ${columns.length} column(s): ${columns.join(", ")}`);
  console.log("Sample values from first row:", JSON.stringify(sampleRow, null, 2), "\n");

  // Build a choice list that includes sample values so the user can orient themselves.
  const columnChoicesWithSkip = [
    ...columns.map(col => ({
      name: `${col}  (e.g. "${String(sampleRow[col] ?? "").slice(0, 60)}")`,
      value: col,
    })),
    { name: "-- Skip / not available --", value: SKIP },
  ];
  const columnChoices = columnChoicesWithSkip.filter(c => c.value !== SKIP);

  // ── Step 2b: fetch workspace members and labels from Linear ──────────────
  const spinner = ora("Fetching workspace members and labels from Linear…").start();
  const client = new LinearClient({ apiKey, apiUrl });
  const [allUsers, allLabels] = await Promise.all([
    client.paginate(client.users, { includeDisabled: false }),
    client.issueLabels(),
  ]);
  spinner.stop();

  const userChoices = [
    { name: "-- Keep CSV value as-is --", value: KEEP },
    ...allUsers.map(u => ({ name: u.email ? `${u.name}  <${u.email}>` : u.name, value: u.name })),
  ];

  const labelChoices = [
    { name: "-- Keep CSV value as-is --", value: KEEP },
    ...(allLabels.nodes ?? []).map(l => ({ name: l.name, value: l.name })),
  ];

  // ── Step 3: map CSV columns to Linear fields ─────────────────────────────
  const { titleCol } = await inquirer.prompt([
    {
      type: "list",
      name: "titleCol",
      message: "Which column is the issue title? (required)",
      choices: columnChoices,
    },
  ]);

  const columnMapping: ColumnMapping = { title: titleCol as string };

  const optionalFieldPrompts: { key: keyof Omit<ColumnMapping, "title">; message: string }[] = [
    { key: "description", message: "Which column is the description?" },
    { key: "status", message: "Which column is the status?" },
    { key: "assignee", message: "Which column is the assignee?" },
    { key: "priority", message: "Which column is the priority?" },
    { key: "dueDate", message: "Which column is the due date?" },
    { key: "labels", message: "Which column should become the issue label?" },
    { key: "url", message: "Which column is the original URL / link?" },
  ];

  for (const { key, message } of optionalFieldPrompts) {
    const { col } = await inquirer.prompt([
      {
        type: "list",
        name: "col",
        message,
        choices: columnChoicesWithSkip,
      },
    ]);
    if (col !== SKIP) {
      columnMapping[key] = col as string;
    }
  }

  // ── Step 4: map status values ────────────────────────────────────────────
  const statusMapping: Record<string, IssueStatus> = {};
  if (columnMapping.status) {
    const uniqueStatuses = [...new Set(rows.map(r => r[columnMapping.status!]).filter(Boolean))];
    if (uniqueStatuses.length > 0) {
      console.log(`\nFound ${uniqueStatuses.length} unique status value(s) in column "${columnMapping.status}".`);
      console.log("Map each one to the matching Linear workflow state type:\n");
      for (const statusVal of uniqueStatuses) {
        const { linearStatus } = await inquirer.prompt([
          {
            type: "list",
            name: "linearStatus",
            message: `  "${statusVal}"  →`,
            choices: STATUS_TYPE_CHOICES,
          },
        ]);
        statusMapping[statusVal] = linearStatus as IssueStatus;
      }
    }
  }

  // ── Step 5: map assignee values to workspace members ────────────────────
  const assigneeMapping: Record<string, string> = {};
  if (columnMapping.assignee) {
    const uniqueAssignees = [
      ...new Set(
        rows
          .flatMap(r => (r[columnMapping.assignee!] ?? "").split(";"))
          .map(v => v.trim())
          .filter(Boolean)
      ),
    ];
    if (uniqueAssignees.length > 0) {
      console.log(`\nFound ${uniqueAssignees.length} unique assignee value(s) in column "${columnMapping.assignee}".`);
      console.log("Map each one to a workspace member:\n");
      for (const assigneeVal of uniqueAssignees) {
        const { linearAssignee } = await inquirer.prompt([
          {
            type: "list",
            name: "linearAssignee",
            message: `  "${assigneeVal}"  →`,
            choices: userChoices,
          },
        ]);
        if (linearAssignee !== KEEP) {
          assigneeMapping[assigneeVal] = linearAssignee as string;
        }
      }
    }
  }

  // ── Step 6: map priority values ──────────────────────────────────────────
  const priorityMapping: Record<string, IssuePriority> = {};
  if (columnMapping.priority) {
    const uniquePriorities = [...new Set(rows.map(r => r[columnMapping.priority!]).filter(Boolean))];
    if (uniquePriorities.length > 0) {
      console.log(`\nFound ${uniquePriorities.length} unique priority value(s) in column "${columnMapping.priority}".`);
      console.log("Map each one to the matching Linear priority level:\n");
      for (const priorityVal of uniquePriorities) {
        const { linearPriority } = await inquirer.prompt([
          {
            type: "list",
            name: "linearPriority",
            message: `  "${priorityVal}"  →`,
            choices: PRIORITY_CHOICES,
          },
        ]);
        priorityMapping[priorityVal] = linearPriority as IssuePriority;
      }
    }
  }

  // ── Step 7: map label values to existing workspace labels ────────────────
  const labelMapping: Record<string, string> = {};
  if (columnMapping.labels) {
    const uniqueLabels = [...new Set(rows.map(r => r[columnMapping.labels!]?.trim()).filter(Boolean))];
    if (uniqueLabels.length > 0) {
      console.log(`\nFound ${uniqueLabels.length} unique label value(s) in column "${columnMapping.labels}".`);
      if (labelChoices.length > 1) {
        console.log("Map each one to an existing workspace label (or keep the CSV value to create a new one):\n");
        for (const labelVal of uniqueLabels) {
          const { linearLabel } = await inquirer.prompt([
            {
              type: "list",
              name: "linearLabel",
              message: `  "${labelVal}"  →`,
              choices: labelChoices,
            },
          ]);
          if (linearLabel !== KEEP) {
            labelMapping[labelVal] = linearLabel as string;
          }
        }
      } else {
        console.log("No existing workspace labels found — CSV label values will be created as new labels.\n");
      }
    }
  }

  // ── Step 8: date format ──────────────────────────────────────────────────
  let dateFormat: DateFormat = "auto";
  if (columnMapping.dueDate) {
    const sampleDate = rows.find(r => r[columnMapping.dueDate!])?.[columnMapping.dueDate!];
    const preview = sampleDate ? `  (sample: "${sampleDate}")` : "";
    console.log(`\nSelect the date format used in column "${columnMapping.dueDate}"${preview}:`);
    const { format } = await inquirer.prompt([
      {
        type: "list",
        name: "format",
        message: "Date format:",
        choices: DATE_FORMAT_CHOICES,
      },
    ]);
    dateFormat = format as DateFormat;
  }

  return new MicrosoftListCsvImporter(
    filePath,
    columnMapping,
    statusMapping,
    priorityMapping,
    dateFormat,
    assigneeMapping,
    labelMapping
  );
};
