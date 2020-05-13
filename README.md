# GeneratePropertyGettersSettersFromBackingFieldsInXLVBA
After creating something like this in VBA:

```
Private Type TModel
   BackingFields...
End Type
Private this As TModel
```
It would be nice to not have to write all of the Property Get/Let/Set procedures.
This project has taken a first swipe at doing just that.

There are two ways provided:
   1) Return a string of the Properties
   2) Let the application add it (not done completely or tested)

Currently, the app distinguishes between Set and Let for known Objects.
Any variable type that is not known (user-defined type, enums, etc. are treated
as objects (i.e. includes "Set") which would have to be adjusted by the user.

Improvements are welcomed.
