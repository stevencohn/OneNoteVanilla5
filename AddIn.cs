//************************************************************************************************
// Copyright © 2021 Steven M Cohn. All rights reserved.
//************************************************************************************************                

namespace OneNoteVanilla5
{
	using Extensibility;
	using System;
	using System.Runtime.InteropServices;
	using Microsoft.Office.Core;
	using Microsoft.Office.Interop.OneNote;
	using System.Windows.Forms;
	using System.Xml.Linq;
	using System.Linq;


	[ComVisible(true)]
	[Guid("4D86B2FD-0C2D-4610-8916-DE24C4BB70B5")] // change this!
	[ProgId("OneNoteVanilla5")] // change this!
	public class AddIn : IDTExtensibility2, IRibbonExtensibility
	{

		// IDTExtensibility2...

		public void OnConnection(
			object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
		{
		}

		public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
			// dispose your stuff here
		}

		public void OnAddInsUpdate(ref Array custom)
		{
		}

		public void OnStartupComplete(ref Array custom)
		{
			// intialize your stuff here
		}

		public void OnBeginShutdown(ref Array custom)
		{
			// cleanup your stuff here
		}


		// IRibbonExtensibility...

		public string GetCustomUI(string RibbonID)
		{
			return @"<?xml version=""1.0"" encoding=""utf -8"" ?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabHome"">
        <group id=""vanillaAddInGroup"" label=""Vanilla Add-in"">
          <button id=""helloButton"" imageMso=""HappyFace"" getLabel=""Say Hello!"" onAction=""SayHello""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
		}


		// Ribbon handlers

		public void SayHello(IRibbonControl control)
		{
			var onenote = new ApplicationClass();

			onenote.GetHierarchy(
				onenote.Windows.CurrentWindow.CurrentNotebookId,
				HierarchyScope.hsSelf,
				out var xml,
				XMLSchema.xs2013);

			var root = XElement.Parse(xml);
			var ns = root.GetNamespaceOfPrefix("one");

			var notebook = root.Elements(ns + "Notebook").FirstOrDefault();
			if (notebook != null)
			{
				var name = notebook.Attribute("name").Value;
				MessageBox.Show($"Hello from {name}!");
			}
			else
			{
				MessageBox.Show($"Error finding notebook name, but Hi anyway!");
			}
		}
	}
}
