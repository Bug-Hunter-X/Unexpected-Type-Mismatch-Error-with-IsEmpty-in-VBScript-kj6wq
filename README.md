# Unexpected Type Mismatch Error with IsEmpty in VBScript

This repository demonstrates a subtle issue with the `IsEmpty` function in VBScript when used with strings that might appear empty but aren't actually empty.  This is frequently encountered when dealing with data from external sources or user input.

The core issue lies in the distinction between an empty string ("") and a `Null` value.  While both might visually appear as nothing, `IsEmpty` handles them differently. An unexpected type mismatch may occur when checking against an empty string.