/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface RawRecord {
  id: string; // Add a unique key for list rendering
  province: string;
  store: string;
  bill: string;
  pCode: string;
  pName: string;
  qty: number;
}

export interface ColumnMapping {
  province: string; // e.g. "F"
  store: string;    // e.g. "C"
  bill: string;     // e.g. "I"
  pCode: string;    // e.g. "J"
  pName: string;    // e.g. "K"
  qty: string;      // e.g. "L"
}

export interface PresetFilter {
  provinces: string[]; // List of selected provinces, e.g. ["พิษณุโลก", "สุโขทัย"]
}
