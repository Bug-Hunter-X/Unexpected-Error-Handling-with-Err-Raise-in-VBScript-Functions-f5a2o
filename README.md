# Unexpected Error Handling with Err.Raise in VBScript Functions

This repository demonstrates a potential issue with VBScript's `Err.Raise` function when used within functions and raising a type mismatch error (Error 13). The error handling might behave differently than expected if the error isn't handled properly within the function itself, or if the calling code doesn't handle the error appropriately.  The example highlights the importance of robust error handling and demonstrates a solution to improve error management.

## Bug Description
When calling `MyFunction` with empty parameters, `Err.Raise` is used. Although intended to raise a type mismatch error, the error's behavior might be unpredictable unless properly managed.  The main issue is that the error might not always propagate in the expected way from the function's scope.

## Solution
The solution involves improving error handling within the `MyFunction` and implementing more robust error checking and trapping in the calling code.