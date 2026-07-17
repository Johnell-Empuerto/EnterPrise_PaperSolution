export {
  registerField,
  getFieldComponent,
  getRegisteredTypes,
  type FieldComponentProps,
  type FieldRegistration,
} from "./FieldRegistry";

// Register built-in field components
import { CheckboxField } from "@/components/Runtime/fields/CheckboxField";
import { DateField } from "@/components/Runtime/fields/DateField";
import { NumberField } from "@/components/Runtime/fields/NumberField";
import { SignatureField } from "@/components/Runtime/fields/SignatureField";
import { KeyboardTextField } from "@/components/Runtime/fields/KeyboardTextField";
import { registerField } from "./FieldRegistry";

registerField("number", NumberField);
registerField("checkbox", CheckboxField);
registerField("date", DateField);
registerField("signature", SignatureField);
registerField("KeyboardText", KeyboardTextField);
