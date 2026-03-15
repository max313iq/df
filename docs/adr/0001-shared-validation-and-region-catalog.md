# ADR 0001: Shared validation and region catalog APIs

## Status

Accepted

## Context

The CLI and GUI previously duplicated portions of input validation and region normalization behavior. The GUI also contained a local fallback template object, which could hide real template/configuration problems and produce unexpected runtime behavior.

## Decision

1. Introduce shared module-level validation helpers:
   - `Convert-ToSanitizedString`
   - `Test-NonEmptyString`
   - `Test-NumericRange`
   - `Test-EmailFormat`
   - `Escape-SpecialCharacters`
2. Add `Get-AzureRegionList` as the shared region catalog API:
   - Azure CLI-backed (`az account list-locations`)
   - in-memory cached
   - graceful fallback list for offline/non-auth scenarios
3. Update region validation to optionally check both:
   - discovered regions from account metadata
   - Azure region catalog
4. Remove GUI local template fallback and fail fast on invalid template load.

## Consequences

- Consistent request/input behavior across interfaces.
- Fewer hidden configuration failures in GUI startup.
- Better maintainability through explicit shared contracts.
- Region validation remains non-breaking (warning-driven) to support constrained environments.
