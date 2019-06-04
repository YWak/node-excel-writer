import { ValueType } from "./value";

export const col2alpha = ((cache: string[]) => (col: number): string => {
  if (cache[col] == null) {
    if (col < 26) {
      cache[col] = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[col];
    } else {
      cache[col] = `${col2alpha(Math.floor((col + 1) / 26) - 1)}${col2alpha(
        col % 26
      )}`;
    }
  }
  return cache[col];
})([]);

export const ref = (r: number, c: number) => `${col2alpha(c)}${r + 1}`;

/**
 * セルの定義。valueとformulaはどちらか一方のみ。
 */
export interface StyledCell {
    /** セルに設定する値 */
    value?: string | null;

    /** セルに設定する数式 */
    formula?: string | null;

    /** セルが表す値の型 */
    type: ValueType;

    /** セルに適用されるスタイルの名称 */
    styles: string[];
  }
