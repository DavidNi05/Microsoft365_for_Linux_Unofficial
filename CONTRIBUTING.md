# Contributing to Microsoft 365 for Linux (Unofficial)

Thank you for your interest in improving this community project! We welcome contributions from everyone. Following these guidelines helps ensure a smooth process for all maintainers and contributors.

---

## How Can I Contribute?

### Reporting Bugs
* Check the existing issues log to make sure the bug hasn't already been reported.
* Open a new issue with a clear title and description.
* Include specific details such as your Linux distribution, Flatpak runtime version, and steps to reproduce the error.
* If applicable, attach terminal logs from running `flatpak run com.microsoft365linux.Unofficial`.

### Suggesting Enhancements
* Open an issue describing the feature you would like to see added.
* Explain why this enhancement would be useful to the broader user base.
* If it relates to a specific Microsoft 365 tool (e.g., Teams, OneNote), outline the expected web interaction model.

### Pull Requests (Code Contributions)
1. Fork the repository and create your branch from `main`.
2. Ensure your JavaScript conforms to clean coding standards (asynchronous operations wrapped in robust `try...catch` blocks).
3. If changing layout components inside `index.html`, adhere to the verified Fluent UI design specs.
4. Verify that any code modification preserves the atomic save operations and strict file system masks (`0o700` / `0o600`).
5. Open a Pull Request targeting the `main` branch with a comprehensive summary of your changes.

---

## Code of Conduct
This project relies entirely on mutual respect, transparent collaboration, and high conscientiousness. Please ensure all interactions remain civil, helpful, and focused on enhancing software usability for the community.
