export interface ValidationRule {
  type: "required" | "pattern" | "maxLength" | "number" | "date";
  message: string;
  validate: (value: string | boolean | null) => boolean;
}

export interface ValidatableField {
  id: string;
  cellReference: string;
  dataType: string;
  required: boolean;
  maxLength: number;
  validationPattern?: string | null;
  validationMessage?: string | null;
}

export interface ValidationResult {
  fieldId: string;
  error: string | null;
}

/**
 * Pure field validation logic.
 * Takes a field definition + value and returns validation errors.
 * No side effects, no store access.
 */
export class ValidationService {
  /**
   * Validates a single field against its definition.
   * Returns error string or null if valid.
   */
  static validateField(field: ValidatableField, value: string | boolean | null): string | null {
    const rules = ValidationService.buildRules(field);
    for (const rule of rules) {
      if (!rule.validate(value)) {
        return rule.message;
      }
    }
    return null;
  }

  /**
   * Validates multiple fields.
   * Returns array of { fieldId, error } results.
   */
  static validateFields(
    fields: ValidatableField[],
    values: Record<string, string | boolean | null>
  ): ValidationResult[] {
    return fields.map((field) => ({
      fieldId: field.id,
      error: ValidationService.validateField(field, values[field.id] ?? null),
    }));
  }

  private static buildRules(field: ValidatableField): ValidationRule[] {
    const rules: ValidationRule[] = [];

    if (field.required) {
      rules.push({
        type: "required",
        message: field.validationMessage ?? `${field.cellReference} is required`,
        validate: (value) => {
          if (value === null || value === undefined) return false;
          if (typeof value === "string" && value.trim().length === 0) return false;
          return true;
        },
      });
    }

    if (field.validationPattern) {
      rules.push({
        type: "pattern",
        message: field.validationMessage ?? `Invalid format for ${field.cellReference}`,
        validate: (value) => {
          if (value === null || value === undefined) return true;
          if (typeof value !== "string" || value.length === 0) return true;
          if (!field.validationPattern) return true;
          try {
            return new RegExp(field.validationPattern).test(value);
          } catch {
            return true;
          }
        },
      });
    }

    if (field.maxLength > 0) {
      rules.push({
        type: "maxLength",
        message: `Maximum ${field.maxLength} characters`,
        validate: (value) => {
          if (value === null || value === undefined) return true;
          if (typeof value !== "string") return true;
          return value.length <= field.maxLength;
        },
      });
    }

    if (field.dataType === "number") {
      rules.push({
        type: "number",
        message: `Must be a valid number`,
        validate: (value) => {
          if (value === null || value === undefined) return true;
          if (typeof value !== "string" || value.length === 0) return true;
          return !isNaN(Number(value));
        },
      });
    }

    if (field.dataType === "date") {
      rules.push({
        type: "date",
        message: `Must be a valid date`,
        validate: (value) => {
          if (value === null || value === undefined) return true;
          if (typeof value !== "string" || value.length === 0) return true;
          return !isNaN(Date.parse(value));
        },
      });
    }

    return rules;
  }
}
