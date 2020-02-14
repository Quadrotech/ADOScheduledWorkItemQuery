import tl = require('azure-pipelines-task-lib/task');
import * as witif from "azure-devops-node-api/interfaces/WorkItemTrackingInterfaces";

export function sortWorkItems(result: witif.WorkItemQueryResult, workItems: witif.WorkItem[]): witif.WorkItem[] {
    if (result.sortColumns != null && result.sortColumns.length > 0) {
        const sortColumn = result.sortColumns[0]!;
        const field = sortColumn.field!;
        const referenceName = field.referenceName!;
        tl.debug("Ordering by " + referenceName + (sortColumn.descending! ? " descending" : " ascending"));
        workItems.sort((a, b) => {
            tl.debug("Processing field [" + referenceName + "]");
            let fieldValueA = a.fields![referenceName];
            let fieldValueB = b.fields![referenceName];
            if (fieldValueA === undefined) {
                return 1;
            }
            if (fieldValueB === undefined) {
                return -1;
            }
            if (typeof (fieldValueA) === "string" && typeof (fieldValueB) === "string") {
                return fieldValueA.localeCompare(fieldValueB);
            }
            else if (typeof (fieldValueA) === "number" && (typeof (fieldValueB) === "number")) {
                return fieldValueA - fieldValueB;
            }
            else if (typeof (fieldValueA) === "object" && fieldValueA.hasOwnProperty("displayName")) {
                return fieldValueA.displayName.localeCompare(fieldValueB.displayName);
            }
            else {
                return fieldValueA > fieldValueB ? 1 : -1;
            }
        });
        if (sortColumn.descending) {
            workItems = workItems.reverse();
        }
    }
    return workItems;
}
