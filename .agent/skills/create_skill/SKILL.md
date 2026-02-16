---
name: create_skill
description: Automates the creation of new skills for the agent, ensuring the correct directory structure and file format.
---

# Create Skill Skill

This skill allows the agent to create new skills in the correct location (`.agent/skills/[skill_name]/SKILL.md`), ensuring strict adherence to the required directory structure.

## Usage

When the user asks to "create a skill" or "make a new capability", follow these steps:

1.  **Determine the Skill Name**: Ask the user for the skill name (e.g., `git_helper`, `pdf_merger`). It should be in snake_case.
2.  **Determine the Description**: Ask what the skill does.
3.  **Run the file creation**:
    *   **Step 1**: Create the directory `.agent/skills/[skill_name]`.
    *   **Step 2**: Create the file `.agent/skills/[skill_name]/SKILL.md`.
    *   **Step 3**: Write the template content with the user's description.

## Template

Use this format for the new `SKILL.md`:

```markdown
---
name: [skill_name]
description: [short description]
---

# [Skill Name Title]

[Detailed description of what the skill does]

## Usage

[Instructions on how the agent should use this skill]
```

## Example Action

If the user says: "Create a skill called 'hello_world' that greets the user."

You should:
1.  Run `mkdir .agent/skills/hello_world`
2.  Create `.agent/skills/hello_world/SKILL.md` with:

```markdown
---
name: hello_world
description: Greets the user.
---

# Hello World Skill
...
```
