# <img src=".pics/lexx_headshot_clear.png" width="100px"/> Lexxeous's VBA Date Compare: <img src=".pics/vb_logo.png" width="100"/>

## Summary:

Simple Visual Basic for Applications project. The VBScipt will be a macro attached to a Microsoft **Excel** document. Upon opening the `.xlsm` file, there is a single reminder event attached and the public `at_time()` procedure is called. The `at_time()` procedure is also recursive and calls itself to persist the attached reminder event as long as the file is open. A loop is used to look through the first 1000 rows of two different sheets. The script then compares today's date (and today's date plus 3 days) with the date values provided in the **Excel** sheets. When there is a match, a message box pop up is generated to inform the user about the reminder.

> For more information you can look at [Microsoft's Visual Basic documentation](https://docs.microsoft.com/en-us/dotnet/visual-basic/).