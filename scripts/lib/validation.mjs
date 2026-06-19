export function createValidation(label) {
  const errors = [];
  return {
    expect(condition, message) {
      if (!condition) {
        errors.push(message);
      }
    },
    fail(message) {
      errors.push(message);
    },
    finish(successMessage) {
      if (errors.length > 0) {
        console.error(`${label} failed.`);
        console.error(errors.map((error) => `- ${error}`).join("\n"));
        process.exit(1);
      }
      console.log(successMessage);
    },
    get errors() {
      return errors;
    }
  };
}
