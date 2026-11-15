using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("Jameel Copilot")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("Jameel Copilot")]
[assembly: AssemblyCopyright("Copyright ©  2026")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("78af2871-0ceb-4451-b80d-455552e37c91")]

// The version attributes are written by the build (JcStampVersion in the
// project file), from the one number the MSI also carries.

// Phase 3a: BuildRestrictFilter and other internals are exposed to the
// xUnit test assembly so we can unit-test DASL filter construction
// without a real Outlook OOM.
[assembly: InternalsVisibleTo("JameelCopilot.Tests")]

