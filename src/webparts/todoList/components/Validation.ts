const validate = (inputs) => {
  //Title errors
  const errors = { title: "", status: "" };
  if (!inputs.title) {
    errors.title = "Title is required";
  }

  //Status Errors
  if (!inputs.status) {
    errors.status = "Status is required";
  }
  return errors;
};

export default validate;
