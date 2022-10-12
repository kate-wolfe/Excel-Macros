Attribute VB_Name = "Guide"
'This is a a guide to how the macros work for the Title Paging List.
'This assumes that you're running this at Main.
'
'When you click one of the buttons, it runs the corresponding 'Display' Macro.
'There are three: BothDisplay, GrayDisplay, and Everything.
'"Both" does local and branch holds, Gray does Gray Bins, and Everything does Everything.
'All three of these macros are towards the bottom of the "Base" module.
'
'All three of these macros activate the "CombinedSort" macro, which starts at the top of the "Base" module.
'This long macro formats the information from the email so that it can then be sorted based on both item location and pickup location.
'This macro also calls the "Stats" Module to record the paging stats.
'
'CombinedSort then calls all of the macros in the "Main" module.
'These macros are the ones that actually sort the list by item location.
'There is one for each section in the building.
'Each one checks the hidden Excel sheet creatively named "Secret" to determine what items belong on their specific list.
'They then filter the full list looking for these items and moving them to their specific sheets once they're found.
'They then check all of the items for their pickup location and adds them to the appropriate sheet.
'From there each sheet is formatted by the "Headers" macro in the Base module.
'This makes the sheets look readable enough to print.
'Finally, the macro unhides the sheet so it is now visible to the user.
'These macros run through one at a time, so the New items are filtered first, then Mezz items, and so on.





'Here's some more info on how to do some more specific things with the paging list.

'TROUBLESHOOTING:
'If you encounter an error while running it, you can click "Debug" to see which line of code is giving you problems.
'Oftentimes the line of code is one of the ones in the "Main" module.
'For example, if it's a problem with an item in the Mezzanine, then Main.MezzList may be the problem.
'The easiest way to deal with this is to close Excel and open it again.
'Then, go to the "Base" module, find the line that calls the relevant macro from "Main", and "comment it out"
'To comment something out, simply put an apostraphe at the beginning of that line of code.
'Excel will think it's a comment rather than a line of code and skip over it.
'This means that the list won't print that section of materials for the day, but hopefully the rest of the list works.




'MOVING AN ITEM'S LOCATION:
'Sometimes a type of item will be relocated to another part of the library.
'For example, when we moved Mystery and Scifi to the Stone Building.
'When this happens, you'll want to update the Paging List so that those items are included on the list for the correct section.
'Item locations are all assigned in a hidden Excel sheet creatively named "Secret".

'To see it, go to File ->Format->Hide and Unhide->Unhide Sheets-> Secret.
'In this sheet there is a column with each section of the library along with a list of the materials that are paged there.
'To move an item from one section to another, simply cut the appropriate cell from one column and paste it into the other.
'Then drag up the other cells in the original column so there are no empty cells below the header.






'SETTING A BRANCH TO BE OPEN OR CLOSED FOR PAGING:
'If a branch is closed, we don't need to page for items going there.
'Likewise, once a closed branch reopens, we want to go back to paging for items going there.

'To toggle this, go to the "Base" module and scroll down a little till you see the below comment:

'"'The next long chunk of code labels CAM and non-CAM Holds so they can be sorted later."
'"These need to be sorted differently for each branch, so there's a nested if statement for each of them."
'"Here's what these mean:
'"1 = Local Pickup"
'"2 = Branch Pickup"
'"3 = Gray bin Pickup"
'"4 = Don't Page (for closed branches). Items with a 4 will not appear on the final list."


'The first block of code there looks like this:
''Main
'If loc = 1 Then
'        If InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/Pickup") = 1 Then
'            Comp.Cells(i, 8).Value2 = 1
'        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/BOUDREAU/Pickup") = 1 Then
'            Comp.Cells(i, 8).Value2 = 4
'        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/CENT SQ/Pickup") = 1 Then
'            Comp.Cells(i, 8).Value2 = 2
'        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/COLLINS/Pickup") = 1 Then
'            Comp.Cells(i, 8).Value2 = 4
'        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/OCONNELL/Pickup") = 1 Then
'            Comp.Cells(i, 8).Value2 = 2
'        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/ONEILL/Pickup") = 1 Then
'            Comp.Cells(i, 8).Value2 = 2
'        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/VALENTE/Pickup") = 1 Then
'            Comp.Cells(i, 8).Value2 = 2
'        Else
'            Comp.Cells(i, 8).Value2 = 3
'        End If




'To change any of the branches from open to closed, simply change its corresponding line of code:
'"'            Comp.Cells(i, 8).Value2 = 2"
'...to replace the last 2 with a 4, like so:
'"            Comp.Cells(i, 8).Value2 = 4"

'For example, if we were going to do this to O'Neill, its line would look like this:
'        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/ONEILL/Pickup") = 1 Then
'            Comp.Cells(i, 8).Value2 = 4


'If a Branch is reopened, change the 4 back to a 2 again.
'You only need to make these changes to the first block of code in this section, the one labelled "Main" shown above.
'The other branches each have a mostly-identical block of code, but you can leave them as they are.


