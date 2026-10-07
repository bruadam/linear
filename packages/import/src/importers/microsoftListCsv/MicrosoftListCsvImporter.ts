import csv from "csvtojson";
import type { Importer, ImportResult, IssuePriority, IssueStatus } from "../../types.ts";

/** Maps each Linear issue field to the corresponding CSV column name. */
export interface ColumnMapping {
  /** CSV column for the issue title (required). */
  title: string;
  /** CSV column for the issue description. */
  description?: string;
  /** CSV column for the issue status. */
  status?: string;
  /** CSV column for the assignee (semicolon-separated multiple values are supported). */
  assignee?: string;
  /** CSV column for the priority. */
  priority?: string;
  /** CSV column for the due date. */
  dueDate?: string;
  /** CSV column whose value should be applied as an issue label. */
  labels?: string;
  /** CSV column for the original issue URL. */
  url?: string;
}

/** Supported explicit date formats. "auto" attempts common formats in order. */
export type DateFormat = "auto" | "DD-MM-YYYY" | "MM-DD-YYYY" | "YYYY-MM-DD" | "DD/MM/YYYY" | "MM/DD/YYYY";

/**
 * Import issues from a Microsoft List CSV export using a user-supplied column
 * and value mapping so the importer is not tied to any specific locale or
 * list configuration.
 */
export class MicrosoftListCsvImporter implements Importer {
  public constructor(
    filePath: string,
    columnMapping: ColumnMapping,
    statusMapping: Record<string, IssueStatus>,
    priorityMapping: Record<string, IssuePriority>,
    dateFormat: DateFormat = "auto",
    /** Maps a raw CSV assignee value to the Linear display name or email used for lookup. */
    assigneeMapping: Record<string, string> = {},
    /** Maps a raw CSV label value to the exact existing Linear label name. */
    labelMapping: Record<string, string> = {}
  ) {
    this.filePath = filePath;
    this.columnMapping = columnMapping;
    this.statusMapping = statusMapping;
    this.priorityMapping = priorityMapping;
    this.dateFormat = dateFormat;
    this.assigneeMapping = assigneeMapping;
    this.labelMapping = labelMapping;
  }

  public get name(): string {
    return "Microsoft List (CSV)";
  }

  public get defaultTeamName(): string {
    return "Microsoft List";
  }

  public import = async (): Promise<ImportResult> => {
    const data = (await csv().fromFile(this.filePath)) as Record<string, string>[];
    const cm = this.columnMapping;

    const importData: ImportResult = {
      issues: [],
      labels: {},
      users: {},
      statuses: {},
    };

    // Collect unique assignees (semicolon-separated values are supported).
    // The resolved name is the mapped value when a mapping exists, otherwise the raw value.
    if (cm.assignee) {
      const allAssignees = new Set<string>();
      for (const row of data) {
        const val = row[cm.assignee];
        if (val) {
          for (const assignee of val.split(";")) {
            const trimmed = assignee.trim();
            if (trimmed) {
              allAssignees.add(this.assigneeMapping[trimmed] ?? trimmed);
            }
          }
        }
      }
      for (const user of allAssignees) {
        importData.users[user] = { name: user };
      }
    }

    // Register statuses from the user-supplied mapping.
    for (const [statusName, statusType] of Object.entries(this.statusMapping)) {
      importData.statuses![statusName] = { name: statusName, type: statusType };
    }

    for (const row of data) {
      const title = cm.title ? row[cm.title] : undefined;
      if (!title) {
        continue;
      }

      // Use only the first assignee when multiple are listed, applying any user mapping.
      const rawFirstAssignee = cm.assignee ? row[cm.assignee]?.split(";")?.[0]?.trim() : undefined;
      const firstAssignee = rawFirstAssignee ? (this.assigneeMapping[rawFirstAssignee] ?? rawFirstAssignee) : undefined;

      const rawDate = cm.dueDate ? row[cm.dueDate] : undefined;
      const dueDate = rawDate ? parseDate(rawDate, this.dateFormat) : undefined;

      const labels: string[] = [];
      if (cm.labels) {
        const rawLabelVal = row[cm.labels]?.trim();
        if (rawLabelVal) {
          // Apply label mapping: use the mapped name if present, otherwise the raw CSV value.
          labels.push(this.labelMapping[rawLabelVal] ?? rawLabelVal);
        }
      }

      const rawStatus = cm.status ? row[cm.status] : undefined;
      const rawPriority = cm.priority ? row[cm.priority] : undefined;

      importData.issues.push({
        title,
        description: cm.description ? row[cm.description] || undefined : undefined,
        status: rawStatus || undefined,
        priority: rawPriority !== undefined ? (this.priorityMapping[rawPriority] ?? 0) : undefined,
        assigneeId: firstAssignee || undefined,
        dueDate,
        labels,
        url: cm.url ? row[cm.url] || undefined : undefined,
      });

      for (const label of labels) {
        if (!importData.labels[label]) {
          importData.labels[label] = { name: label };
        }
      }
    }

    return importData;
  };

  private filePath: string;
  private columnMapping: ColumnMapping;
  private statusMapping: Record<string, IssueStatus>;
  private priorityMapping: Record<string, IssuePriority>;
  private dateFormat: DateFormat;
  private assigneeMapping: Record<string, string>;
  private labelMapping: Record<string, string>;
}

/**
 * Parse a date string using the given format.
 * When format is "auto", ISO (YYYY-MM-DD) is tried first, then DD-MM-YYYY,
 * then native Date parsing as a last resort.
 */
function parseDate(input: string, format: DateFormat): Date | undefined {
  if (!input) {
    return undefined;
  }

  const tryDate = (isoString: string): Date | undefined => {
    const d = new Date(isoString);
    return isNaN(d.getTime()) ? undefined : d;
  };

  if (format === "YYYY-MM-DD" || format === "auto") {
    const m = input.match(/^(\d{4})[-/](\d{2})[-/](\d{2})$/);
    if (m) {
      return tryDate(`${m[1]}-${m[2]}-${m[3]}`);
    }
  }

  if (format === "DD-MM-YYYY" || (format === "auto" && /^\d{2}-\d{2}-\d{4}$/.test(input))) {
    const m = input.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (m) {
      return tryDate(`${m[3]}-${m[2]}-${m[1]}`);
    }
  }

  if (format === "DD/MM/YYYY") {
    const m = input.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) {
      return tryDate(`${m[3]}-${m[2]}-${m[1]}`);
    }
  }

  if (format === "MM-DD-YYYY") {
    const m = input.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (m) {
      return tryDate(`${m[3]}-${m[1]}-${m[2]}`);
    }
  }

  if (format === "MM/DD/YYYY") {
    const m = input.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) {
      return tryDate(`${m[3]}-${m[1]}-${m[2]}`);
    }
  }

  // Last resort: native Date parsing.
  return tryDate(input);
}
