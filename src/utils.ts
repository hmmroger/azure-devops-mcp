// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export const apiVersion = "7.2-preview.1";
export const batchApiVersion = "5.0";

export function getEnumKeys<T extends Record<string, string | number>>(enumObject: T): string[] {
  return Object.keys(enumObject).filter((key) => isNaN(Number(key)));
}
