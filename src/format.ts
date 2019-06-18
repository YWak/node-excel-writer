const builtin = (): NumberFormat[] => [
  { id: 0, formatCode: "General" },
  { id: 1, formatCode: "0" },
  { id: 2, formatCode: "0.00" },
  { id: 3, formatCode: "#,##0" },
  { id: 4, formatCode: "#,##0.00" },
  { id: 9, formatCode: "0%" },
  { id: 10, formatCode: "0.00%" },
  { id: 11, formatCode: "0.00E+00" },
  { id: 12, formatCode: "# ?/?" },
  { id: 13, formatCode: "# ??/??" },
  { id: 14, formatCode: "mm-dd-yy" },
  { id: 15, formatCode: "d-mmm-yy" },
  { id: 16, formatCode: "d-mmm" },
  { id: 17, formatCode: "mmm-yy" },
  { id: 18, formatCode: "h:mm AM/PM" },
  { id: 19, formatCode: "h:mm:ss AM/PM" },
  { id: 20, formatCode: "h:mm" },
  { id: 21, formatCode: "h:mm:ss" },
  { id: 22, formatCode: "m/d/yy h:mm" },
  { id: 37, formatCode: "#,##0 ;(#,##0)" },
  { id: 38, formatCode: "#,##0 ;[Red](#,##0)" },
  { id: 39, formatCode: "#,##0.00;(#,##0.00)" },
  { id: 40, formatCode: "#,##0.00;[Red](#,##0.00)" },
  { id: 45, formatCode: "mm:ss" },
  { id: 46, formatCode: "[h]:mm:ss" },
  { id: 47, formatCode: "mmss.0" },
  { id: 48, formatCode: "##0.0E+0" },
  { id: 49, formatCode: "@" }
];

export class Format {

    private builtins: FormatDictionary = {};

    private customs: FormatDictionary = {};

    public constructor(locale?: string) {
        for (const format of builtin()) {
            this.add(this.builtins, format);
        }

        if (locale != null) {
            const formats = require(`./locale/${locale}`).formats() as NumberFormat[];
            for (const format of formats) {
                this.add(this.builtins, format);
            }
        }
    }

    private add(dict: FormatDictionary, format: NumberFormat) {
        if (dict[format.formatCode] == null) {
            dict[format.formatCode] = format.id;
        }
    }

    public getId(formatCode: string) {
        let id;

        if ((id = this.builtins[formatCode]) != null) {
            return id;
        }

        if ((id = this.customs[formatCode]) != null) {
            return id;
        }

        id = (this.builtins.length + this.customs.length + 1);
        this.customs[formatCode] = id;

        return id;
    }
}

export type FormatDictionary = {[pattern: string]: number};

export interface NumberFormat {
    id: number;
    formatCode: string;
}
