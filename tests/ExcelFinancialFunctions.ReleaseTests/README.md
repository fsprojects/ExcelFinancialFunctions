# Release Tests

This project tests new release candidate nuget packages after they are deployed to the NuGet Gallery. It should always include a reference to the latest released version of the package on NuGet Gallery.

When a new release candidate package is uploaded, we can then update this project to refer to it, then run these tests. This helps ensure that we didn't break anything in the upgrade process.

## Differences

This project has some differences from the unit tests or the interop tests.

* Runs against .NET Framework 4.6.1. This is the oldest supported configuration.
* Written in C#. Helps ensure that the library functions correctly in that language.
* Uses the latest package from NuGet Gallery. Not the locally built code.