
"use strict";

function OnNewShellUI( shellUI ) {
	/// <summary>The entry point of ShellUI module.</summary>
	/// <param name="shellUI" type="MFiles.ShellUI">The new shell UI object.</param> 

	// Register to listen new shell frame creation event.
	shellUI.Events.Register( Event_NewShellFrame, newShellFrameHandler );
}

function newShellFrameHandler( shellFrame ) {
	/// <summary>Handles the OnNewShellFrame event.</summary>
	/// <param name="shellFrame" type="MFiles.ShellFrame">The new shell frame object.</param> 

	// Register to listen the started event.
	shellFrame.Events.Register( Event_Started, getShellFrameStartedHandler( shellFrame ) );
}

function getShellFrameStartedHandler( shellFrame ) {
	/// <summary>Gets a function to handle the Started event for shell frame.</summary>
	/// <param name="shellFrame" type="MFiles.ShellFrame">The current shell frame object.</param> 
	/// <returns type="MFiles.Events.OnStarted">The event handler.</returns>

	// Return the handler function for Started event.
	return function() {
		var vault = shellFrame.shellUI.Vault;
			var LoggedinUser= shellFrame.ShellUI.Vault.CurrentLoggedInUserID				
	var S,A=0;
		//var vault = shellFrame.shellUI.Vault;
		
    var cnt = shellFrame.ShellUI.Vault.SessionInfo.UserAndGroupMemberships

    for (i = 2; i <= cnt.length; i++) { 
    S = shellFrame.ShellUI.Vault.SessionInfo.UserAndGroupMemberships.Item(i).UserOrGroupID
        //shellFrame.ShowMessage(S);
	if(S==289){
        A = 1;
    //    shellFrame.ShowMessage("It has been set to 1");
        }
}
    if(A!=1) {
		//shellFrame.ShowMessage("No :" + A);
	}else{
    	// Shell frame object is now started.
		
		// Create some commands.
		var commandShow1 = shellFrame.Commands.CreateCustomCommand( "Merge Document" );
		var commandShow2 = shellFrame.Commands.CreateCustomCommand( "Create Board Agenda" );

		// Set command icons.
		shellFrame.Commands.SetIconFromPath( commandShow1, "png/tennis_ball.ico" );
		shellFrame.Commands.SetIconFromPath( commandShow2, "png/flower_red.ico" );
		
		// Add a command to the context menu.
		shellFrame.Commands.AddCustomCommandToMenu( commandShow1, MenuLocation_ContextMenu_Bottom, 0 );
		
		// Add a commands to the task pane.
		shellFrame.TaskPane.AddCustomCommandToGroup( commandShow1, TaskPaneGroup_Main, -101 );
		shellFrame.TaskPane.AddCustomCommandToGroup( commandShow2, TaskPaneGroup_Main, -100 );
		
		// Set the command handler function.
		shellFrame.Commands.Events.Register( Event_CustomCommand, function( command ) {
		
			// Branch by command.
			if( command == commandShow1 ) {
					
				var vault = shellFrame.shellUI.Vault;
				var BoardAgendaClass;
				var selectedItems = shellFrame.ActiveListing.CurrentSelection;
				var flag = 0;
				for( var i = 1; i <= selectedItems.ObjectVersions.Count; ++i )
					{	
						var objectversion = selectedItems.ObjectVersions.Item(i).ObjVer;
						var oProp = shellFrame.ShellUI.Vault.ObjectPropertyOperations.GetProperties(objectversion);
						BoardAgendaClass = oProp.SearchForProperty(100).TypedValue.GetValueAsLocalizedText() //get the object type 
						if (BoardAgendaClass == "Board Agenda Temp") {
						var BoardAgendaNamePropertyAlias = vault.PropertyDefOperations.GetPropertyDefIDByAlias("Board Agenda Name")
						var MatterName= oProp.SearchForProperty(BoardAgendaNamePropertyAlias).TypedValue.GetValueAsLocalizedText()
						//shellFrame.ShowMessage(BoardAgendaClass);
						}
					else{
						flag = 1;
						shellFrame.ShowMessage("Select a Board Agenda");

					}
				}

				if (BoardAgendaClass == "Board Agenda Temp") {
						var oShell = new ActiveXObject("Shell.Application");
						//this is for the server \\\\hqtwi275\\DocumentMerge\\DocumentMerge.exe
						var commandtoRun = "\\\\DESKTOP-A8GUMF6\\DocumentMerge\\DocumentMerge.exe"
						
						//OnNewDashboard(dashboard)
						oShell.ShellExecute(commandtoRun,MatterName,"","open","1");
					}
					else if(flag == 0){
						shellFrame.ShowMessage("Select a Board Agenda");

					}
				// Show a message.				
			} else if(command == commandShow2) {
				
					var vault = shellFrame.shellUI.Vault;
				    var selectedItems = shellFrame.ActiveListing.CurrentSelection;
				    var MatterClass;
				    var fla = 0;
				for( var i = 1; i <= selectedItems.ObjectVersions.Count; ++i )
					{	
						var objectversion = selectedItems.ObjectVersions.Item(i).ObjVer;
						 var oProp = shellFrame.ShellUI.Vault.ObjectPropertyOperations.GetProperties(objectversion); //get the object version
						 MatterClass = oProp.SearchForProperty(100).TypedValue.GetValueAsLocalizedText() //get the object type
							 if (MatterClass == "Matter") {
							 var MatterAtomPropertyAlias = vault.PropertyDefOperations.GetPropertyDefIDByAlias("Matter Atom") // find the property id by alias
							 var MatterAtom = oProp.SearchForProperty(MatterAtomPropertyAlias).TypedValue.GetValueAsLocalizedText() // get the matter atom
						 }
						 else{
						 	fla = 1
							shellFrame.ShowMessage("You must select the Matter");
						 }
						  						
					}

					if (MatterClass == "Matter") {

						var oShell = new ActiveXObject("Shell.Application");
					
						var commandtoRun = "\\\\DESKTOP-A8GUMF6\\DocumentMerge\\MFilesTemplate.exe"
						
						//OnNewDashboard(dashboard)
						oShell.ShellExecute(commandtoRun,MatterAtom,"","open","1");
					}

					else if(fla == 0){
						shellFrame.ShowMessage("Select a Matter");


					}
			}
		} );
}//End If 
	};
}
