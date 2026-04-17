import csv from "csvtojson";
import type { Importer, ImportResult, IssuePriority, IssueStatus } from "../../types.ts";

type MicrosoftListPriority = "Presserende" | "Vigitig" | "Mellem" | "Lav";
type MicrosoftListStatus = "Backlog" | "Klar til start" | "I gang" | "Afventer" | "Udført" | string;

interface MicrosoftListIssueType {
  Titel: string;
  Beskrivelse: string;
  Tildelt: string;
  Forfallsdato: string;
  Status: MicrosoftListStatus;
  Prioritet: MicrosoftListPriority;
  Faggruppe: string;
}

/**
 * Import issues from a Microsoft List CSV export.
 *
 * @param filePath  path to csv file
 */
export class MicrosoftListCsvImporter implements Importer {
  public constructor(filePath: string) {
    this.filePath = filePath;
  }

  public get name(): string {
    return "Microsoft List (CSV)";
  }

  public get defaultTeamName(): string {
    return "Microsoft List";
  }

  public import = async (): Promise<ImportResult> => {
    const data = (await csv().fromFile(this.filePath)) as MicrosoftListIssueType[];

    const importData: ImportResult = {
      issues: [],
      labels: {},
      users: {},
      statuses: {},
    };

    // Collect unique assignees (Tildelt can be semicolon-separated)
    const allAssignees = new Set<string>();
    for (const row of data) {
      if (row.Tildelt) {
        for (const assignee of row.Tildelt.split(";")) {
          const trimmed = assignee.trim();
          if (trimmed) {
            allAssignees.add(trimmed);
          }
        }
      }
    }

    for (const user of allAssignees) {
      importData.users[user] = { name: user };
    }

    // Register statuses
    const statusTypes: Record<string, IssueStatus> = {
      Backlog: "backlog",
      "Klar til start": "unstarted",
      "I gang": "started",
      Afventer: "unstarted",
      Udført: "completed",
    };

    for (const [statusName, statusType] of Object.entries(statusTypes)) {
      importData.statuses![statusName] = { name: statusName, type: statusType };
    }

    for (const row of data) {
      const title = row.Titel;
      if (!title) {
        continue;
      }

      // Use the first assignee when multiple are listed
      const firstAssignee = row.Tildelt?.split(";")[0]?.trim();
      const assigneeId = firstAssignee || undefined;

      const dueDate = row.Forfallsdato ? parseDanishDate(row.Forfallsdato) : undefined;

      const labels: string[] = [];
      if (row.Faggruppe?.trim()) {
        labels.push(row.Faggruppe.trim());
      }

      importData.issues.push({
        title,
        description: row.Beskrivelse || undefined,
        status: row.Status || undefined,
        priority: mapPriority(row.Prioritet),
        assigneeId,
        dueDate,
        labels,
      });

      for (const label of labels) {
        if (!importData.labels[label]) {
          importData.labels[label] = { name: label };
        }
      }
    }

    return importData;
  };

  // -- Private interface

  private filePath: string;
}

/**
 * Parse a date string in DD-MM-YYYY format (as used in Norwegian/Danish Microsoft List exports).
 */
function parseDanishDate(input: string): Date | undefined {
  const match = input.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!match) {
    return undefined;
  }
  const [, day, month, year] = match;
  return new Date(`${year}-${month}-${day}`);
}

const mapPriority = (input: MicrosoftListPriority): IssuePriority => {
  const priorityMap: { [k: string]: IssuePriority } = {
    Presserende: 1,
    Vigitig: 2,
    Mellem: 3,
    Lav: 4,
  };
  return priorityMap[input] ?? 0;
};
