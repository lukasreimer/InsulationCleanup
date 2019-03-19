# InsulationCleanup
Revit addin for cleaning up insulation

## Description
This addin provides a command for deleting all rogue insulation (insulation which are on a different workset than the hosting objects).
It is done by filtering through all insulation elements and then deleting and recreating the ones on non-matching worksets.
The command can also be used just for inspection without automatically changing the elements.
