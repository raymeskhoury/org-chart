// Utility functions

export function ASSERT(condition: boolean, message?: string): void {
  if (!condition) {
    if (message !== null) {
      throw new Error(message);
    }
    throw new Error("Assertion failed");
  }
}

export class Util {
  // Remove the given element from the given array.
  static removeFromArray(array: any[], element: any): void {
    const index = array.indexOf(element);
    if (index > -1) {
      array.splice(index, 1);
    }
  }

  static async appendChildAndWait(parent: Node, child: Node): Promise<void> {
    const waitForAppend = new Promise<void>(resolve => {
      const observer = new MutationObserver(
        (records: MutationRecord[], observer: MutationObserver) => {
          for (const record of records) {
            for (const node of record.addedNodes) {
              if (node === child) {
                observer.disconnect();
                resolve();
                return;
              }
            }
          }
        }
      );
      observer.observe(parent, {childList: true});
    });
    parent.appendChild(child);
    return waitForAppend;
  }
}
