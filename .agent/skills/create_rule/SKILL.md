---
name: create_rule
description: Automates the creation of global rules for the agent, ensuring they are placed in the correct directory.
---

# Create Rule Skill

This skill allows the agent to create new rule files in the correct location (`.agent/rules/`), preventing directory structure errors.

## Usage

When the user asks to "create a rule" or "add a global rule", follow these steps:

1.  **Determine the Rule Name**: Ask the user what the rule is about if not specified (e.g., `terminology`, `coding_style`).
2.  **Determine the Content**: confirm the content of the rule.
3.  **Run the file creation**:
    *   Target Directory: `.agent/rules/`
    *   File Name: `[rule_name].md`
    *   **CRITICAL**: ensure the directory `.agent/rules` exists before writing.

## Example Action

If the user says: "Create a rule to always use snake_case for Python variables."

You should:
1.  Create `.agent/rules/python_style.md` (or similar name).
2.  Write the rule content into that file.

```markdown
# Python Style Rules

- **Variable Naming**: Always use `snake_case` for variable names.
```

## Setup Verification

This skill does not require external libraries. It relies on the agent's ability to write files.
