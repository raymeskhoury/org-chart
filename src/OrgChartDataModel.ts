import {action, autorun, makeObservable, observable} from "mobx";
import * as Papa from "papaparse";

const COLUMN_NAMES = ["id", "parentId", "name", "position", "color"];

function toString(object: any): string {
  if (object === undefined || object === null) {
    return "";
  }
  return String(object);
}

export class OrgChartEntry {
  @observable
  id: string;

  @observable
  parentId: string;

  @observable
  name: string;

  @observable
  position: string;

  @observable
  color: string;

  constructor(id = "", parentId = "", name = "", position = "", color = "") {
    this.id = id;
    this.parentId = parentId;
    this.name = name;
    this.position = position;
    this.color = color;
    makeObservable(this);
  }
}

export class OrgChartDataModel {
  @observable
  data: OrgChartEntry[];

  constructor() {
    this.data = [];
    makeObservable(this);
    this.resetToSample();

    const storedData = localStorage.getItem("data");
    if (storedData !== null) {
      console.error(storedData);
      this.fromCsv(storedData);
    }

    autorun(() => {
      localStorage.setItem("data", this.getCsv());
    });
  }

  @action
  resetToSample(): void {
    this.data = [
      new OrgChartEntry("1", "", "Bianca Toscano", "Director", "#ffffff"),
      new OrgChartEntry(
        "2",
        "1",
        "Aasa Andrejev",
        "Manager, Marketing",
        "#0d2747"
      ),
      new OrgChartEntry(
        "3",
        "1",
        "Paul Lohmus",
        "Manager, Products",
        "#0d2747"
      ),
      new OrgChartEntry(
        "4",
        "2",
        "Sergio Udinese",
        "PR Coordinator",
        "#97daff"
      ),
      new OrgChartEntry(
        "5",
        "2",
        "Mattia Sabbatini",
        "Content Strategist",
        "#97daff"
      ),
      new OrgChartEntry("6", "3", "Mai Aare", "Engineering Lead", "#ffffff"),
      new OrgChartEntry("7", "3", "Aet Kangro", "Design Lead", "#0d2747"),
      new OrgChartEntry("8", "4", "Aili Mihhailov", "PR Specialist", "#97daff"),
      new OrgChartEntry("9", "4", "Lemme Kangur", "PR Assistant", "#97daff"),
      new OrgChartEntry("10", "5", "Alice Cattaneo", "Copywriter", "#97daff"),
      new OrgChartEntry(
        "11",
        "6",
        "Helbe Piip",
        "Software Engineer",
        "#0d2747"
      ),
      new OrgChartEntry("12", "6", "Riccardo Buccho", "Intern", "#0d2747"),
      new OrgChartEntry("13", "7", "Jana Piip", "UX Designer", "#97daff"),
    ];
  }

  getCsv(): string {
    const result = Papa.unparse(this.data, {
      newline: "\n",
    });
    return result;
  }

  @action
  fromCsv(text: string): string {
    const result = Papa.parse(text, {header: true, skipEmptyLines: true});
    if (result.errors.length > 0) {
      return (
        "Error: " +
        (result.errors[0].row !== undefined
          ? `row ${result.errors[0].row}: `
          : "") +
        result.errors[0].message
      );
    }
    for (const field of result.meta.fields!) {
      const lowerColumnNames = COLUMN_NAMES.map(val => val.toLowerCase());
      if (!lowerColumnNames.includes(field.toLowerCase())) {
        return `Error: Invalid column name "${field}`;
      }
    }
    const data = [];
    for (const row of result.data as any[]) {
      const entry = new OrgChartEntry();
      entry.id = toString(row.id);
      entry.parentId = toString(row.parentId);
      entry.name = toString(row.name);
      entry.position = toString(row.position);
      entry.color = toString(row.color);
      data.push(entry);
    }

    this.data = data;
    return "";
  }
}
