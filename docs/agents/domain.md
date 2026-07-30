# Domain docs

This repository uses a single-context domain documentation layout.

## Before exploring

Read these sources when they exist and are relevant:

- `CONTEXT.md` at the repository root for the shared domain language.
- ADRs under `docs/adr/` for established technical decisions.

If either source does not exist, proceed without treating its absence as an error. Skills such as `/domain-modeling`, `/grill-with-docs`, and `/improve-codebase-architecture` may create them when the project actually resolves new terminology or architectural decisions.

## Layout

```text
/
|- CONTEXT.md
|- docs/
|  `- adr/
`- application source files
```

## Vocabulary

Use terms defined in `CONTEXT.md` consistently in code, tests, issues, and design documents. If a needed term is missing, check whether the codebase already uses another name before adding it to the glossary.

## Architecture decisions

Read relevant ADRs before changing an established design. If proposed work conflicts with an ADR, state the conflict explicitly instead of silently overriding the prior decision.
