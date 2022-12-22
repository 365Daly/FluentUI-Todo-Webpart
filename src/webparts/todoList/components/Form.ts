import { useState } from "react";
import { ITodoFormValues } from "./Interface";

const useForm = (initialValues: ITodoFormValues, validate) => {
  const [inputs, setInputs] = useState<ITodoFormValues>(initialValues);
  const [errors, setErrors] = useState<ITodoFormValues>({
    title: "",
    status: "",
  });

    const handleSubmit = (fn: (values) => void) => {
       return (e) => {
         e.preventDefault();
         const validationErrors = validate(inputs);
         const noErrors = Object.keys(validationErrors).length === 0;
         setErrors(validationErrors);
         if (noErrors) {
           fn(inputs);
         } else {
           console.error("errors try again", validationErrors);
         }
       };
  };

  const handleInputChange = (field: string, value: string | number): void => {
    setInputs((inputs) => ({ ...inputs, [field]: value }));
  };

  return {
    handleSubmit,
    handleInputChange,
    inputs,
    errors,
  };
};

export default useForm;
