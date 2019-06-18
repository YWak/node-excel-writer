/**
 * セルのスタイルを表します。
 */
export interface Style {

    /** フォントのスタイル */
    font?: FontStyle;

    /** 罫線のスタイル */
    border?: BorderStyle;

    /** 塗りつぶしのスタイル */
    fill?: FillStyle;

    /** 値の表示スタイル */
    format?: string | number;

    /** 値の横方向の配置スタイル */
    align?: HolizontalAlignment;

    /** 値の縦方向の配置スタイル */
    verticalAlign?: VerticalAlignment;
}

/**
 * フォントのスタイルを表します。
 */
export interface FontStyle {

    /** 文字の大きさ */
    size?: number;

    /** フォント名 */
    name?: string;

    /** 文字色 */
    color?: string;

    /** 太文字かどうか */
    bold?: boolean;

    /** イタリック体かどうか */
    italic?: boolean;

    /** 打ち消し線を引くかどうか */
    strike?: boolean;

    /** 下線を引くかどうか */
    underline?: boolean;
}

/**
 * 罫線のスタイルを表します。
 */
export interface BorderStyle {

    /** セル上側の罫線 */
    top?: LineStyle;

    /** セル右側の罫線 */
    right?: LineStyle;

    /** セル下側の罫線 */
    bottom?: LineStyle;

    /** セル左側の罫線 */
    left?: LineStyle;
}

/**
 * 線のスタイルを表します。
 */
export interface LineStyle {

    /** 線の形状 */
    style?: LineType;

    /** 線の色 */
    color?: string;
}

/**
 * セルの塗りつぶしのスタイルを表します。
 */
export interface FillStyle {
    /** 塗りつぶし色 */
    color?: string;

    /** 塗りつぶしパターン */
    pattern?: string;
}

/**
 * セルの適用条件を表す関数です。
 *
 * @param row 0を起点とした、セルの行
 * @param col 0を起点とした、セルの列
 * @returns セルに適用するかどうか
 */
export type Predicate = (row: number, col: number) => boolean;

/** 横方向の文字列の配置スタイル */
export type HolizontalAlignment = "left" | "center" | "right";

/** 縦方向の文字列の配置スタイル */
export type VerticalAlignment = "top" | "center" | "bottom";

/** 罫線のスタイル */
export type LineType =
    | "none"
    | "thin"
    | "medium"
    | "dashed"
    | "dotted"
    | "thick"
    | "double"
    | "hair"
    | "medium-dashed"
    | "dash-dot"
    | "medium-dash-dot"
    | "dash-dot-dot"
    | "medium-dash-dot-dot"
    | "slanted-dash-dot";

export class StyleManager {

    private styles: { [keys: string]: Style } = {};

    private rules: Rule[] = [];

    private cache: string[] = [];

    public addStyle(name: string, style: Style) {
        if (this.styles[name]) {
            throw new Error(`style '${name}' is already defined`);
        }

        this.styles[name] = style;
    }

    public resetRules() {
        this.rules = [];
    }

    public addRule(name: string, predicate: Predicate) {
        if (!this.styles[name]) {
            throw new Error(`style '${name}' is not defined`);
        }

        this.rules.push({ name, predicate });
    }

    public getStyleIndex(row: number, col: number) {
        const names: string[] = [];

        for (const rule of this.rules) {
            if (rule.predicate(row, col) && !names.find(n => n == rule.name)) {
                names.push(rule.name);
            }
        }

        const key = names.join(",")

        for (let i = 0; i < this.cache.length; i++) {
            if (this.cache[i] === key) {
                return i;
            }
        }
        this.cache.push(key);
        return this.cache.length - 1;
    }
}

export const mergeStyles = (styles: Style[]) => {

};

interface Rule {
    name: string;
    predicate: Predicate;
}
