# Skill: Interface Layout & Design Token Boundaries

## Operational Constraints
1. Every design layout suggestion MUST adhere strictly to a deterministic 8px spacing grid ($8\text{px}, 16\text{px}, 24\text{px}, 32\text{px}$). No arbitrary margin or padding values are permitted.
2. Maintain accessible typographic hierarchy: Title (24px/Bold), Subtitle (18px/Medium), Body (14px/Regular).
3. Component Selection Protocol: When prompting for structural shifts, invoke `ui_layouts_library` to check for predefined accessible flexbox/grid wrapper frameworks before writing custom layouts.
4. Color Token Rules: Use semantic variables exclusively (`var(--background)`, `var(--foreground)`, `var(--primary)`) to preserve seamless light/dark mode adaptation.