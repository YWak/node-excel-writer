export interface ValueConverter<T> {
    readonly name: string;
    readonly type: ValueType;
    canApply(obj: unknown): obj is T;
    convert(value: T): string | null;
}

export type ValueType = "number" | "date" | "boolean" | "string" | "blank";

export abstract class ValueConverterSkelton<T> implements ValueConverter<T> {
    private _type: ValueType;
    private _name: string;

    public constructor(type: ValueType, name: string) {
        this._type = type;
        this._name = name;
    }

    get type(): ValueType { return this._type; }
    get name(): string { return this._name; }

    public abstract canApply(obj: unknown): obj is T;
    public abstract convert(value: T): string | null;
}

export class NumberValueConverter extends ValueConverterSkelton<number> {
    public constructor() {
        super("number", "number");
    }

    public canApply(obj: unknown): obj is number {
        return typeof obj === "number" && Number.isFinite(obj);
    }

    public convert(value: number): string {
        return value.toString();
    }
}

export class DateValueConverter extends ValueConverterSkelton<Date> {
    public constructor() {
        super("date", "date");
    }

    public canApply(obj: unknown): obj is Date {
        return typeof obj === "object" && obj instanceof Date;
    }
    public convert(value: Date): string {
        return value.toISOString();
    }
}

export class NullConverter extends ValueConverterSkelton<null> {
    public constructor() {
        super("blank", "null");
    }

    public canApply(obj: unknown): obj is null {
        return typeof obj == null;
    }
    public convert(value: null): null {
        return null;
    }
}

export class AnyValueConverter extends ValueConverterSkelton<any> {
    public constructor() {
        super("string", "any");
    }

    public canApply(obj: unknown): obj is any {
        return true;
    }
    public convert(value: any): string {
        if (value == null) {
            return "";
        } else {
            return value.toString();
        }
    }
}
