Late Binding: VBScript allows late binding, meaning you can use objects without explicitly declaring their type.  While flexible, this can lead to runtime errors if an object or method doesn't exist.  The error messages can be cryptic and difficult to trace.  Example:
```vbscript
Set obj = CreateObject("NonExistent.Object")
' Error occurs here if NonExistent.Object doesn't exist
```