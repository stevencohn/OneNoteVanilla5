# OneNoteVanilla5
Demo a vanilla OneNote add-in using .NET 5.0

https://github.com/dotnet/core-setup/blob/master/Documentation/design-docs/COM-activation.md

* .NET Framework used mscoree.dll as a shim to bootrap COM activation
* .NET Core uses comhost.dll and one is generated for the project; see Install.reg


# How-to

## Setup

Requirements: Visual Studio 16.8 or higher, .NET 5.0 SDK

### Create Project

1. Visual Studio -> Create a new project

2. C#/Windows/Library -> Choose Class Library

3. Enter names

4. Target Framework -> .NET 5.0

5. Edit the csproj file:

   Change the TargetFramework to net5.0-windows and enable COM hosting

     ```xml
     <TargetFramework>net5.0-windows</TargetFramework>
     <EnableComHosting>true</EnableComHosting>
     ```

   To use Windows Forms, add the UseWindowsForms element

     <UseWindowsForm>true</UseWindowsForms>


### Add Dependencies

6. Add COM reference Microsoft.Office 16.0 Object Library  
   This adds Interop.Microsoft.Office.Core

   a. Note that some people have found it necessary to enable _Embed Interop Type_ in the
      dependency proeprties but it seems to work without it.

7. Add COM reference Microsoft OneNote 15.0 Type Library  
   This adds Interop.Microsoft.Office.Interop.OneNote

   a. Note that some people have found it necessary to enable _Embed Interop Type_ in the
      dependency proeprties but it seems to work without it.

8. Browse to reference <VSpath>\Common7\IDE\PublicAssemblies\extensibility.dll  
   This adds Extensiblity


### Create AddIn class

9. Create a new class (or rename the default Class1.cs) for the addin  
   The name can be anything you want

10. Add these using statements

    ```csharp
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.OneNote;
    ```

11. Add the following attributes to the addin class

    ```csharp
    [ComVisible(true)]
    [Guid("4D86B2FD-0C2D-4610-8916-DE24C4BB70B5")]
    [ProgId("OneNoteVanilla5")]
    ```

    replacing the Guid with your own unique Guid  
    replacing the ProgId with a unique ID; this will be recorded in the System Registry

12. Extend the class with these interfaces:

    ```csharp
    : IDTExtensibility2     // adds lifetime handlers to the addin
    : IRibbonExtensibility  // adds ribbon handlers to the addin
    ```
