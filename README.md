# Cryptic Runtime Errors Due to Late Binding in VBScript

This repository demonstrates a common, yet often cryptic, error in VBScript stemming from its support for late binding. Late binding, while offering flexibility, can lead to runtime errors that are difficult to diagnose due to unclear error messages.

The `bug.vbs` file showcases the problematic code, while `bugSolution.vbs` provides a corrected version with early binding for improved error handling.

## Problem

VBScript's late binding allows you to use objects without explicit type declarations.  However, this flexibility comes at the cost of runtime errors if an object or method is not found. The resulting error messages might not pinpoint the exact issue, making debugging challenging.

## Solution

The solution involves using early binding.  By declaring the object type explicitly, the VBScript interpreter can check for the existence of the object and its members at compile time, avoiding runtime errors.  This leads to more robust and predictable code.