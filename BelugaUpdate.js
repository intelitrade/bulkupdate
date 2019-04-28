//debugger;
//debugger;
function updateFile()
{
	try{
		//debugger;
		//debugger;
		
		WScript.Echo("Welcome To Project Beluga (Bulk Update...)");
		WScript.Echo("");
		WScript.Echo("Brought To You By Lucy Stevens Pty [Group IT]");
		WScript.Echo("");		
		WScript.Echo("Contacts");
		WScript.Echo("");
		WScript.Echo("WhatsApp +27 64 676 7320 / +27 72 045 6865");
		WScript.Echo("");			
		WScript.Echo("Process Started...");
		WScript.Echo("");
		//debugger;
		//debugger;
		var oCaseWareApp = WScript.CreateObject("CaseWare.Application");
		var sProgramPath = sCaseWareDir = oCaseWareApp.ApplicationInfo("ProgramPath");

		var oShell = WScript.CreateObject("WScript.Shell");
		
		var sFilePath = oShell.CurrentDirectory +"\\Beluga.txt";

		oShell.Run("cmd /c dir/s/b *.ac? > Beluga.txt");
		//Destroy the shell object no longer needed
		oShell = null;
		
		WScript.Echo("File Discovery");
		WScript.Echo("Locating CaseWare Files...");
		WScript.Echo("");
		WScript.Sleep(20000); //'Sleeps for 20 seconds - Give DOS enough time to create the input file (file with list of caseware files)

		if(!checkIfFileExist(sFilePath) || !checkIfFolderExists(sCaseWareDir))
		{
			if(!checkIfFileExist(sFilePath))
				WScript.Echo("Input File Not Found");
			
			if(!checkIfFolderExists(sCaseWareDir))
				WScript.Echo("CaseWare Folder Not Found");
			
			return;
		}
		var sFileData = readTextFile(sFilePath);
		//debugger;
		if(sFileData!="")
		{
			//Get the list of files to update
			var aFilesToUpdate = sFileData.split("\n");
			if(isInputValid(aFilesToUpdate) && aFilesToUpdate.length>0)
			{				
				aFilesToUpdate = removeBlanksElementsFromArray(aFilesToUpdate);
				
				var iLength = aFilesToUpdate.length;
				WScript.Echo("CaseWare Client Files Found "+iLength);
				WScript.Echo("");
				var iFilesUpdated = 0;
				
				for(var i=0;i<iLength;i++)
				{
					//Put a try catch so that if one file fails to update the rest can continue
					try{
						var sFile = aFilesToUpdate[i];
						sFile = sFile.replace("\r","");
						sFile = sFile.replace("\n","");
						if(!isInputValid(sFile))
							continue;
						
						if(checkIfFileExist(sFile))
						{
							//var sTempName = sFile.replace(".ac_",".ac")
							var sFileExtension = getFileExtension(sFile);
							if(sFileExtension=="ac")
								sTempName = sFile.replace(".ac",".ac_");
							else
								sTempName = sFile.replace(".ac_",".ac");
							
							//Check if compressed and uncompressed file exist
							if(checkIfFileExist(sFile) && checkIfFileExist(sTempName))
							{
								WScript.Echo("Compressed & Uncompressed File In Same Directory: "+sFile);
								WScript.Echo("");
								WScript.Echo("This Directory Will Not Be Processed");
								WScript.Echo("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx");
								WScript.Echo("");							
								continue;
							}


							WScript.Echo("Checking For Product Updates");
							WScript.Echo("");
							
							var today = new Date();
							var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
							var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
							var dateTime = date+' '+time;								
							
							WScript.Echo("Start Time: "+dateTime);
							WScript.Echo("");	
							
							var vReturn = checkForProductUpdates(sFile,oCaseWareApp,sCaseWareDir);
							
							var today = new Date();
							var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
							var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
							var dateTime = date+' '+time;	
							
							WScript.Echo("End Time: "+dateTime);
							WScript.Echo("");
							
							//debugger;
							//debugger;
							
							if(!isInputValid(vReturn) || vReturn=="No Update" || vReturn=="No update")
							{
								WScript.Echo("No Updates For File: "+sFile);
								WScript.Echo("");
								continue;
							}
							bCompressFileAfter = false;
							
							if(getFileExtension(sFile)=="ac_")
							{
								
								WScript.Echo("Uncompressing File: "+sFile);
								WScript.Echo("");
								var today = new Date();
								var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
								var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
								var dateTime = date+' '+time;
								WScript.Echo("Start Time: "+dateTime);
								WScript.Echo("");	
								//uncompress the file
								oCaseWareApp.Clients.Uncompress(sFile);
								sFile = sFile.replace(".ac_",".ac");
								WScript.Echo("File uncompressed: "+sFile);
								WScript.Echo("");
								var today = new Date();
								var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
								var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
								var dateTime = date+' '+time;
								WScript.Echo("End Time: "+dateTime);
								WScript.Echo("");
								
								bCompressFileAfter = true;
							}else{
								sFile = sFile.replace(".ac_",".ac");
							}
							
							//convert file
							WScript.Echo("Converting File: "+sFile);
							
							var today = new Date();
							var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
							var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
							var dateTime = date+' '+time;
							
							WScript.Echo("Start Time: "+dateTime);
							WScript.Echo("");
							
							oCaseWareApp.Clients.Convert(sFile);
							
							WScript.Echo("Process Completed For File: "+sFile);
							WScript.Echo("");
							
							var today = new Date();
							var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
							var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
							var dateTime = date+' '+time;
							
							WScript.Echo("End Time: "+dateTime);
							WScript.Echo("");

							var iUpdates = vReturn.length;
							var aToRun = [];
							var sUpdateFunction = "";
							
							WScript.Echo(iUpdates+" Update(s) Available For File: "+sFile);
							WScript.Echo("");								
							
							for(var j=0;j<iUpdates;j++)
							{
								var aUpdate = vReturn[j];
								/*var iUpdate = aUpdate.length;
								for(var k=0;k<iUpdate;k++)
								{*/
								var sUpdate = aUpdate[0];
								var aUpdateFunction = sUpdate.split("_");
								sUpdateFunction = aUpdateFunction[0];
								if(sUpdateFunction=="updateClientFile")
								{
							
									WScript.Echo("Copy Template Started: "+sFile);
									WScript.Echo("");
									
									var today = new Date();
									var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
									var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
									var dateTime = date+' '+time;
								
									WScript.Echo("Start Time: "+dateTime);
									WScript.Echo("");
									
									performCopy(oCaseWareApp,sFile,sCaseWareDir);
									WScript.Echo("Copy template complete: "+sFile);
									WScript.Echo("");
									
									var today = new Date();
									var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
									var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
									var dateTime = date+' '+time;	
									
									WScript.Echo("End Time: "+dateTime);
									WScript.Echo("");
						
									aToRun[aToRun.length] = sUpdateFunction;
								}
								if(sUpdateFunction == "updateProbeClientFile1" || sUpdateFunction == "updateProbeClientFile"){
									
									aToRun[aToRun.length] = sUpdateFunction;//"updateProbeClientFile";
								}
								if(sUpdateFunction == "convertTaxClientFile"){
									aToRun[aToRun.length] = sUpdateFunction;
								}								
							}
							
							if(sUpdateFunction == "updateProbeClientFile1")
							{
								var s_CVWFileToExecuteAgainst = sProgramPath+"Scripts\\SA IFRS\\Remote Scripts\\UPDATE.cvw";
								sUpdateFunction = updateProbeClientFile;
							}else{
								var s_CVWFileToExecuteAgainst = sFile.replace(".ac","")+"FSNG0000ZAFS.cvw";
							}
							
							var sScriptName = sProgramPath + "Scripts\\SA IFRS\\CQS_PatchLib.scp"
							if(sUpdateFunction == "convertTaxClientFile"||sUpdateFunction == "updateProbeClientFile1"||sUpdateFunction=="updateClientFile"){
								var sFunctionName = "AdaptITUpdatesToRun";
								var sParam = aToRun;
							}else{
								var sFunctionName = sUpdateFunction;
								var sParam = "";
							}	
							
							var s_ClientPath = CQSGetFilePathLib(sFile);
							var s_ClientName = CQSGetFileNameOnlyLib(sFile);
							
							var today = new Date();
							var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
							var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
							var dateTime = date+' '+time;
							
							WScript.Echo("Updating File: "+sFile);
							WScript.Echo("");
							WScript.Echo("Start Time: "+dateTime);
							WScript.Echo("");
							
							oCaseWareApp.CVConvert.ConvertOneCaseViewScript(sProgramPath,s_ClientPath, s_ClientName, s_CVWFileToExecuteAgainst, sScriptName,  sFunctionName, sParam, 0);
							
							WScript.Echo("File Update Complete: "+sFile);
							WScript.Echo("");
							
							var today = new Date();
							var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
							var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
							var dateTime = date+' '+time;						
							WScript.Echo("End Time: "+dateTime);
							WScript.Echo("");
							
							iFilesUpdated++;
							//}
							//Check if the file needs to be compressed - always put the file back in the state it was found, it also reduces the footprint on the clients machine
							//Will be handled by another application
							if(bCompressFileAfter == true){
								/*var today = new Date();
								var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
								var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
								var dateTime = date+' '+time;
								
								WScript.Echo("Compressing file: "+sFile);
								WScript.Echo("");
								WScript.Echo("Start Time: "+dateTime);
								WScript.Echo("");
								
								try{
									oCaseWareApp.Clients.Compress(sFile,0);
								}catch(e)
								{
									WScript.Echo("File compression failed: "+sFile);
									WScript.Echo("");								
								}
								var today = new Date();
								var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
								var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
								var dateTime = date+' '+time;						
								WScript.Echo("End time: "+dateTime);
								WScript.Echo("");*/
								
							}
							
						}else{
							WScript.Echo("File Not Found: "+sFile);
							WScript.Echo("------------------------------------------------------------");
						}
					}catch(e)
					{
						WScript.Echo("Error With File: "+sFile);
						WScript.Echo("");
						WScript.Echo("Error Description: "+e.description);
						WScript.Echo("");
					}
				}
				
				WScript.Echo("");
				WScript.Echo(iFilesUpdated+" / "+iLength+" Files Updated");
				WScript.Echo("");
			}
		}else{
			WScript.Echo("Nofication: The Input File Is Empty");
			WScript.Echo("");
		}
		
		WScript.Echo("Thank You For Using Project Beluga");
		WScript.Echo("");
		WScript.Echo("Bulk Update Process Completed...");
		WScript.Echo("");
		WScript.Echo("");
		WScript.Echo("Upcoming projects"); 
		WScript.Echo("");
		WScript.Echo("Beluga XL");
		WScript.Echo("Update files in parallel");
		WScript.Echo("");
		WScript.Echo("Beluga Notification");
		WScript.Echo("Get process notifications");
		WScript.Echo("");
		WScript.Echo("Beluga Next (Planned release - October / November 2019)");
		WScript.Echo("Run updates Online & other processes");
		WScript.Echo("");
		
		/*var oShell = WScript.CreateObject("WScript.Shell");
		oShell.Popup("Thank You For Using Project Beluga\n\n"+iFilesUpdated+" / "+iLength+" Files Updated\n\nUpcoming projects\n\nBeluga XL\nUpdate files in parallel\n\nBeluga Notification\nGet process notifications\n\nBeluga Online\nRun updates online & other processes");//, 5, "Beluga");*/
		
		//debugger;
	}catch(e)
	{
		WScript.Echo("Error: "+e.description);
		WScript.Echo("");
	}finally{
		today = null;
		oCaseWareApp = null;
		aToRun = null;
		aUpdate = null;
		vReturn = null;
	}
}

function removeBlanksElementsFromArray(aArray)
{
	try{
		if(isInputValid(aArray))
		{
			var iLength = aArray.length;
			for(var i=0;i<iLength;i++)
			{
				if(aArray[i]==""){
					aArray.splice(i,1);
				}
			}
		}
		return aArray;
	}catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}finally{
		
	}
}

function unCompressCasewareFile(sCQSMasterFilePath,oCaseWareApp)
{
	try
	{
		if(!oCaseWareApp)
			var oCaseWareApp =  new ActiveXObject("CaseWare.Application");
		
		//check if the file is compressed
		if(IsFileCompressed(sCQSMasterFilePath))
		{
			oCaseWareApp.Clients.Uncompress(sCQSMasterFilePath);
		}
	}
	catch(e)
	{
		WScript.Echo("Error: "+e.description);
		WScript.Echo("");
	}
}
function performCopy(oCaseWareApp,sFile,sCaseWareDir)
{
	//debugger
	try
	{
		//get the source file path
		var sCQSMasterFilePath = GetTemplatePath("", "FULL_IFRS", "2.8", 0, 2, "","","",oCaseWareApp,sFile);
		var sProduct = getMetaData("CWCustomProperty.Product",sCQSMasterFilePath,oCaseWareApp);
		unCompressCasewareFile(sCQSMasterFilePath,oCaseWareApp);
		var sClientFilePath = sFile;//curDoc.FileName;
	
		if(!oCaseWareApp)
			var oCaseWareApp = new ActiveXObject("CaseWare.Application");
		
		if(oCaseWareApp) {
			l_copytemplate = oCaseWareApp.CopyTemplate;
			
			l_copytemplate.SourceFilePath =  sCaseWareDir+ "\\Scripts\\SA IFRS\\Remote Scripts\\"+sProduct+"\\";//application.CaseView.ProgramPath + "Scripts\\SA IFRS\\Remote Scripts\\"+sProduct+"\\"; //CQSGetFilePathLib(sCQSMasterFilePath) + "\\"
			l_copytemplate.SourceFileName = sProduct+".ac"; //CQSGetFileNameOnlyLib(sCQSMasterFilePath) + ".ac"
			
			//set the destination path
			l_copytemplate.DestinationFilePath = CQSGetFilePathLib(sClientFilePath) + "\\"
			l_copytemplate.DestinationFileName = CQSGetFileNameOnlyLib(sClientFilePath) + ".ac"
			
			//copy some components
			l_copytemplate.CopyAll = false;
			var docs = l_copytemplate.CopyDocuments;
			docs.SelectAllCopyFlag (false);
							
			l_copytemplate.CopyTrialBalance = 0;
			l_copytemplate.CopyTaxCodes = 0;
			l_copytemplate.CopySecurity = 0;
			l_copytemplate.CopyUserDefinedInfo = 0;
			l_copytemplate.CopyUnits = 0;
			l_copytemplate.CopySplitUpAccounts = 0;
			l_copytemplate.CopyJournalTypes = 0;
			l_copytemplate.CopyTickmarks = 0;
			l_copytemplate.CopyTaxonomy = 0;
			l_copytemplate.CopyHistorySettings = 0;
			//Mo - I suspect the external db values will be copied across using the onFileNew Script attached to each document
			//I must test this
			l_copytemplate.CopyCVExternalData = 0;
			l_copytemplate.ClearAccountBalances = 0;
			l_copytemplate.ClearSpreadsheetAnalysis = 0;
			l_copytemplate.ClearProgramChecklistInfo = 0;
			l_copytemplate.ClearForeignExchange = 0;
			l_copytemplate.ClearRoleCompletion = 0;
			l_copytemplate.ClearProgramAssertionInfo = 0;
			l_copytemplate.ClearAnnotationText = 0;
			l_copytemplate.ClearDocumentReferences = 0;
			l_copytemplate.ClearTickmarks = 0;
			l_copytemplate.ClearAnnotationReferences = 0;
			l_copytemplate.ClearCVDocumentReferences = 0;
			l_copytemplate.ClearCVTickmarks = 0;
			l_copytemplate.ClearCVNotes = 0;
			//CopyTemplate.Silent = true;
			l_copytemplate.Silent = false;
			//debugger
			//debugger
			l_copytemplate.CopyMappings = true;
			l_copytemplate.CopyGroupings(0)=true;
			l_copytemplate.CopyGroupings(1)=true;
			l_copytemplate.CopyGroupings(2)=true;
			l_copytemplate.CopyGroupings(3)=true;
			l_copytemplate.CopyGroupings(4)=true;
			l_copytemplate.CopyGroupings(5)=true;
			l_copytemplate.CopyGroupings(6)=true;
			l_copytemplate.CopyGroupings(7)=true;
			l_copytemplate.CopyGroupings(8)=true;
			l_copytemplate.CopyGroupings(9)=true;
			l_copytemplate.CopyGroupingsTo(1) = 1;
			l_copytemplate.CopyGroupingsTo(2) = 2;
			l_copytemplate.CopyGroupingsTo(3) = 3;
			l_copytemplate.CopyGroupingsTo(4) = 4;
			l_copytemplate.CopyGroupingsTo(5) = 5;
			l_copytemplate.CopyGroupingsTo(6) = 6;
			l_copytemplate.CopyGroupingsTo(7) = 7;
			l_copytemplate.CopyGroupingsTo(8) = 8;
			l_copytemplate.CopyGroupingsTo(9) = 9;
			l_copytemplate.CopyGroupingsTo(10) = 10;
			//debugger
			if (l_copytemplate.CopyAll)
			{l_copytemplate.DoCopyLite();}
			else
			{
				try {
					l_copytemplate.DoCopy();
				} catch(e) {
					//if it comes here the client clicked cancel or there was an error
					//logError(e)
					WScript.Echo("Error: "+e.description);
					WScript.Echo("");
					success = false;
					g_bDoCopy = false;
					canceled = true;
					//debugger;
				}
			}
		}
	}
	catch(e)
	{
		WScript.Echo("Error: "+e.description);
		WScript.Echo("");		
	}finally{
		//Replace specific CaseView documents		
		AdaptITReplaceDocuments(sCQSMasterFilePath, sClientFilePath,1);
		return;
	}
}


function CQSGetFilePathLib(sPathandFileName)
{ 
	try
	{
		var sResult = "";
		var oFSO = new ActiveXObject("Scripting.FileSystemObject");
		if(oFSO)
		{
			sResult = oFSO.GetParentFolderName(sPathandFileName)
		}
	}
	catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}     
	return sResult;
}

function CQSGetFileNameOnlyLib(sPathandFileName)
{

	try
	{
		var sResult = "";
		var oFSO = new ActiveXObject("Scripting.FileSystemObject");
		if(oFSO)
		{
			sResult = oFSO.GetBaseName(sPathandFileName)
		}
	}
	catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}     
	return sResult;
}


function AdaptITReplaceDocuments(sCQSMasterFilePath, sClientFilePath,iMakeBackUpCopy)
{
	try{
		//debugger;
		var oFso = new ActiveXObject("Scripting.FileSystemObject");
		if(oFso)
		{
			//Hard coded array but vould be moved into a text file
			var aDocsToReplace = ["FSNG0000ZAFS.cvw","SI000000ZAFS.cvw","FIRMSET0ZAFS.cvw","PIS00000.cvw","CFWKS000.cvw"];		
			var iDocs = aDocsToReplace.length;
			for(var i=0;i<iDocs;i++)
			{
				var sFile = aDocsToReplace[i];
				var sFileToCopy = CQSGetFilePathLib(sCQSMasterFilePath) + "\\" +CQSGetFileNameOnlyLib(sCQSMasterFilePath)+sFile;
				var sFileToReplace = CQSGetFilePathLib(sClientFilePath) + "\\" +CQSGetFileNameOnlyLib(sClientFilePath)+sFile;
				//check if the file to be copied exists
				if(oFso.FileExists(sFileToCopy))
				{	
					//Check if a backup of the file to be replaced needs to be made
					if(oFso.FileExists(sFileToReplace) && iMakeBackUpCopy===1)
					{
						var sNameOfBackUpCopy = sFileToReplace.replace(".cvw","-Copy.cvw");
						oFso.CopyFile(sFileToReplace,sNameOfBackUpCopy, 1);
						//delete the original
						oFso.DeleteFile(sFileToReplace,1)				
					}
					//copy the required file to the proposed destination
					oFso.CopyFile(sFileToCopy,sFileToReplace, 1);				
				}
			}
		}
	
	}catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}finally{
		oFso = null;
	}
}


function isInputValid(vInput)
{
	try{
		bTrue = true;
		
		if(vInput===""||typeof(vInput)==="undefined"||vInput===null)
			bTrue = false;
		
		return bTrue;		
	}catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}
}


function checkForProductUpdates(sClientFile_Path,oCaseWareApp,sProgramPath)
{
	try{
		var aAvailableUpdates = new Array();
		//var oCWApp = new ActiveXObject("CaseWare.Application");
		//var sClientFile_Path =  oCWApp.ActiveClient.FileName;
		var sCountry = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Country",oCaseWareApp);
		var sProdName = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Product",oCaseWareApp);
		//var sCopyTemplate = Client.CaseViewData.dataGroupFormId("FORMATTING", "CONTROLS", "COPYTEMPLATE");
		//check if the client has the CQS master tempate installed for the product he is currently using
		//if not, exit this function as updates for this product cannot be executed
		//var bExit = shouldUpdatesRun(sProdName,"",sClientFile_Path)

		//new code added for the 2009 templates, from now on all updates will be running through this
		aAvailableUpdates = CQSCheckforAddtionalUpdates(oCaseWareApp, aAvailableUpdates,sClientFile_Path)
		//if there are updates available launch the update html dialogue
		if (aAvailableUpdates.length>0)
		{

			var aHTMLArray = new Array()

			for(var p=0;p<aAvailableUpdates.length;p++)
			{
				if(aAvailableUpdates[p][4]==1)
					aHTMLArray[aHTMLArray.length] = aAvailableUpdates[p];
			}
			//if there are no update options where the display value = 1 then exit the update
			//this will cater for old files that have many updates available but dont have the relevant template installed, and so prevent a blank update wizard from appearing
			if(aHTMLArray.length == 0)
			{
				//return
			}
			//before launching the HTML, create addiditional info to be used inside the HTML
			var oAdditionalInfo = new ActiveXObject("Scripting.Dictionary");
		 
			var bCheckForSMEtemplate = isTemplateInstalled("SME",oCaseWareApp,sClientFile_Path);
			var bCheckForIFRStemplate = isTemplateInstalled("IFRS7",oCaseWareApp,sClientFile_Path);
			if(sProdName!="PROBEATA" && sProdName!="MODIFIEDCASH"){
				var bCheckForRelatedEntitySME = isRelatedEntitySelected("SME",oCaseWareApp,sClientFile_Path);
				var bCheckForRelatedEntityIFRS = isRelatedEntitySelected("IFRS7",oCaseWareApp,sClientFile_Path);
			}
				

			oAdditionalInfo.Add("SMEinstalled",bCheckForSMEtemplate)
			oAdditionalInfo.Add("IFRSinstalled",bCheckForIFRStemplate)
			oAdditionalInfo.Add("RelatedEntitiesSME",bCheckForRelatedEntitySME)
			oAdditionalInfo.Add("RelatedEntitiesIFRS",bCheckForRelatedEntityIFRS)


			oAdditionalInfo.Add("country",sCountry)
			
			if(oAdditionalInfo)
			{
				//var ilaunchCWUpdateDialog = Client.CaseViewData.dataGroupFormId("UPDATE","DIALOG","LAUNCHCWDIALOG");
				var sCurrentProduct =  CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Product",oCaseWareApp);
				var sClientLanguage = getMetaData("CWCustomProperty.ProductLanguage",sClientFile_Path,oCaseWareApp);
				var sCountry = getMetaData("CWCustomProperty.Country",sClientFile_Path,oCaseWareApp);
				var sCode = ""
				var sProbeClientVersion = CQSGetMetaData(sClientFile_Path,"CWCustomProperty.Probe_Cversion",oCaseWareApp)
				var sProduct = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Product",oCaseWareApp)
				var sVersion = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Mapping",oCaseWareApp)
				var sLanguage = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.ProductLanguage",oCaseWareApp)
				var sCountry = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Country",oCaseWareApp)
				var sCQSMasterFilePath = GetTemplatePath(sCode, sProduct, sVersion, 0, sLanguage, sCountry)
				var sTemplateVersion = getMetaData("CWCustomProperty.Mapping",sCQSMasterFilePath,oCaseWareApp);
				var sTemplateProbeVersion = getMetaData("CWCustomProperty.Probe_Cversion",sCQSMasterFilePath,oCaseWareApp);
				//var sDocID = document.identifier;
				if((sCurrentProduct=="FULL_IFRS" && sTemplateVersion>="3.7") || (sProdName=="PROBEATA" && sTemplateProbeVersion>sProbeClientVersion) || (sProdName=="MODIFIEDCASH" && sTemplateVersion>sVersion) && ilaunchCWUpdateDialog == 1)
				{
					//var oUpdatesToRun = launchClientFileUpdateDialog(oCaseWareApp,aHTMLArray,sLanguage,sVersion,sProduct,"","",oAdditionalInfo)
					return aHTMLArray; //{'updatelist':aHTMLArray,'language':sLanguage,'version':sVersion,'product':sProduct,'additionalinfo',oAdditionalInfo};
				}
			}
		}else{
			return "No update";
		}
	}catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}finally{
		
	}
}

function isRelatedEntitySelected(sProduct,oCaseWareApp,sClientFile_Path)
{
	try
	{
        if(!oCaseWareApp)
		{
			var oCaseWareApp = new ActiveXObject("CaseWare.Application");
			var sClientFile_Path =  oCaseWareApp.ActiveClient.FileName;
		}
		//get the client file object
       // var sClientFilePath =sClientFile_Path
		//debugger;
		//debugger;
		//get the entity type currently selected on this client file
		/*var oClientFile = oCaseWareApp.Clients.Open(sClientFile_Path,"SUP","SUP");
		var sClientFileEntityWithPipes = oClientFile.CaseViewData.dataGroupFormId("", "", "AYENTITY");
		*/
		//Work aroudn for the 2 lines above, in order to bypasss security
		//debugger;
		//debugger;
		var sCWSrcDbfFilePath = sClientFile_Path.substr(0,sClientFile_Path.lastIndexOf(".ac"))+"CV.dbf";
		var oSrcSystemCaseViewData = oCaseWareApp.SystemCaseViewData(sCWSrcDbfFilePath);
		var sClientFileEntityWithPipes = "";
		if(oSrcSystemCaseViewData)
			sClientFileEntityWithPipes = oSrcSystemCaseViewData.GetGroupFormIdData("","","AYENTITY",0);
		
		//The entity value in the database gets returned as 4 tokens separated by pipes 000001|-00001|CO|Company
        //get the third token
		 var sClientFileEntity = "";
        if(isInputValid(sClientFileEntityWithPipes)){

            sClientFileEntity = sClientFileEntityWithPipes.split("|")[2]
        }

		if(sProduct=="IFRS for SME") sProduct = "SME";
		if(sProduct=="IFRS") sProduct = "FULL_IFRS";
		if(sProduct=="IFRS7") sProduct = "FULL_IFRS";

		sCode = ""
		sVersion = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Mapping", oCaseWareApp);
		sProductNameInMetaData = sProduct


		if(sVersion == "")
		{
			sVersion = "0.0"
		}


		//get the language of the client file
		if (sLanguage = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.ProductLanguage", oCaseWareApp)==4) //Afrikaans
		{
			var sLanguage = "4"
		}
		else //English
		{
			var sLanguage = "2"
		}

        var sCountry = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Country", oCaseWareApp)
		var sCQSMasterFilePath = GetTemplatePath(sCode, sProductNameInMetaData, sVersion, 0, sLanguage, sCountry)


        //get the entity type related to the master template that is applicable to this client file
        if(sCQSMasterFilePath!=""){
        var sCQSmasterFileEntity = CQSGetMetaData(sCQSMasterFilePath, "CWCustomProperty.RelatedEntities",oCaseWareApp);
        }else{
            return 1;
        }

		//if the meta data custom property CWCustomProperty.RelatedEntities is just a comma, it means that this CQS master template is applicable to all entities selected on the client file
		if(sCQSmasterFileEntity == ",")
		{
			return 1
		}

		//if CWCustomProperty.RelatedEntities does not exist, allow option to show for the time being. Once the install sets have been updated with the latest templates and meta data
		//we can remove this if statement. This if statement will cater for developers who get the latest scripts from sourcsafe but do not yet have the latest templates with the updated meta data
		if(sCQSmasterFileEntity == "")
		{
			return 1
		}

		var aCQSmasterFileEntity = sCQSmasterFileEntity.split(",")
		var bTempEntity = 0
		//check if any of the related entities match the entity selected in the client file
		for (var i = 0; i < aCQSmasterFileEntity.length; i++)
		{
			var sThisEntity = aCQSmasterFileEntity[i]
			if(sThisEntity == sClientFileEntity)
			{
				return 1
			}
		}

		//if no matching entity was found in the above for loop, return 0
		return bTempEntity


	}
	catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}finally{
		oClientFile = null;
		oSrcSystemCaseViewData = null;
	}
}



function verifyCountry(sCountry)
{
//KN - Q4 2011
	try
	{
		//run tests for undefined and " ZA "		
		try
		{
			if(sCountry==null || typeof(sCountry)=="undefined" || sCountry==" " || sCountry.trim()=="")
				sCountry = "ZA"
		}
		catch(e)
		{
			sCountry = "ZA"
			WScript.Echo("Error: "+e.description);
		}

		return sCountry
	}
	catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}
}

String.prototype.trim = string_trim
function string_trim()
{
	try{
		return this.replace(/(^\s*)|(\s*$)/g, "");
	}catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}
}
function PROBEIsanUpdateAvailable(sClientProduct, sClientVersion, sClientLanguage, sCountry,oCaseWareApp,sClientFile_Path)
{
//This function will check if there is an update available for the probe client
//file. if there is it will return a true, if not it will return false
  try
  {
	 if(!oCaseWareApp)
	 {	var oCaseWareApp = new ActiveXObject("CaseWare.Application");
		var sClientFile_Path = oCaseWareApp.ActiveClient.FileName;
	 }
    var bResult = false;
	var sPathandFileName = "";
	var sClientSubProduct = "";
    var sTemplatePath = GetTemplatePath("", sClientProduct, sClientVersion, 0, sClientLanguage, sCountry,sPathandFileName,sClientSubProduct,oCaseWareApp);

	//var sClientFilePath = sClientFile_Path;
	var sProbeClientVersion = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Probe_Cversion",oCaseWareApp);
	var sClientProbeProduct = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Probe_Product",oCaseWareApp);
    if(sTemplatePath!="" && sClientProbeProduct!="")
    {
		var sProbeTemplateVersion = CQSGetMetaData(sTemplatePath, "CWCustomProperty.Probe_Cversion",oCaseWareApp);
		if(sProbeClientVersion < sProbeTemplateVersion)
		{
			bResult = true;
		}
	}

	return bResult;
  }
  catch(e)
  {
    WScript.Echo("Error: "+e.description);
  }
}


function CQSIsanUpdateAvailable(sClientProduct, sClientVersion, sClientLanguage, sClientEntity, sCountry,sPathandFileName,sClientSubProduct,oCaseWareApp,sClientFile_Path)
{
//This function will check if there is an update available for the client
//file. if there is it will return a true, if not it will return false
  try
  {
    var bResult = false;
    var sTemplatePath = GetTemplatePath("", sClientProduct, sClientVersion, 0, sClientLanguage, sCountry,sPathandFileName,sClientSubProduct,oCaseWareApp,sClientFile_Path);
	//sCode, sProduct, sVersion, bAlternate, sLanguage, sCountryParam,sPathandFileName,sClientSubProduct,oCaseWareApp,sClientFile_Path
    if(sTemplatePath!="")
    {
		var sIgnoreCQSupdate = CQSGetMetaData(sTemplatePath, "CWCustomProperty.ignoreCQSupdate",oCaseWareApp);
		if(sIgnoreCQSupdate==1)
			return false

		var sTemplateVersion = CQSGetMetaData(sTemplatePath, "CWCustomProperty.Mapping",oCaseWareApp);
		if(sClientVersion < sTemplateVersion)
		{
            bResult = true;
		}
	}

	return bResult;
  }
  catch(e)
  {
    WScript.Echo("Error: "+e.description);
  }

}

function CQSCheckforAddtionalUpdates(oCaseWareApp, aAvailableUpdates,sClientFile_Path)
{
	////debugger
	//this function will receive an array that will be passed in from
	//"checkForProductUpdates". we will take the array, check for new versions
	//and then add to the array if there are new updates available
	//If there are updates available then we will change the properties of
	//the other items not to display
	try
	{	 
		var aCurrentUpdates = new Array();
		if (aCurrentUpdates.length>0)
		{
			for(var p=0;p<aCurrentUpdates.length;p++)
			{
				if(aCurrentUpdates[p][3]=="U")
					aCurrentUpdates[p][4] = 0;
			}
		}
		//debugger;
		//debugger;
		var sClientProduct = getMetaData("CWCustomProperty.Product",sClientFile_Path,oCaseWareApp);
		var sClientVersion =  getMetaData("CWCustomProperty.Mapping",sClientFile_Path,oCaseWareApp);
		var sClientLanguage = getMetaData("CWCustomProperty.ProductLanguage",sClientFile_Path,oCaseWareApp);
		if(sClientProduct!="PROBEATA" && sClientProduct!="MODIFIEDCASH"){
			//debugger;
			//debugger;
			/*var oClientFile = oCaseWareApp.Clients.Open(sClientFile_Path,"SUP","SUP");
			var sClientEntity = oClientFile.CaseViewData.dataGroupFormId("", "", "AYENTITY").split("|")[2];*/
			var sCWSrcDbfFilePath = sClientFile_Path.substr(0,sClientFile_Path.lastIndexOf(".ac"))+"CV.dbf";
			var oSrcSystemCaseViewData = oCaseWareApp.SystemCaseViewData(sCWSrcDbfFilePath);
			var sClientEntity = "";
			
			if(oSrcSystemCaseViewData){
				var sClientFileEntityWithPipes = oSrcSystemCaseViewData.GetGroupFormIdData("","","AYENTITY",0);
				if(isInputValid(sClientFileEntityWithPipes))
				{
					sClientEntity = sClientFileEntityWithPipes.split("|")[2];
					if(!isInputValid(sClientEntity))
						sClientEntity = "";
				}
			}
		}
		var sCountry = getMetaData("CWCustomProperty.Country",sClientFile_Path,oCaseWareApp);
		var sProbeClientVersion = getMetaData("CWCustomProperty.Probe_Cversion",sClientFile_Path,oCaseWareApp);
		var sClientSubProduct = "";
		var sPathandFileName = "";
		var bUpdateAvailable = CQSIsanUpdateAvailable(sClientProduct, sClientVersion, sClientLanguage, sClientEntity, sCountry,sPathandFileName,sClientSubProduct,oCaseWareApp,sClientFile_Path);
		if(bUpdateAvailable)
		{
			var sTemplatePath = GetTemplatePath("", sClientProduct, sClientVersion, 0, sClientLanguage, sCountry,sPathandFileName,sClientSubProduct,oCaseWareApp,sClientFile_Path);
			var sTemplateProduct =  CQSGetMetaData(sTemplatePath, "CWCustomProperty.Product",oCaseWareApp);
			var sTemplateVersion = CQSGetMetaData(sTemplatePath, "CWCustomProperty.Mapping",oCaseWareApp);
			var sTemplateDescr = CQSGetMetaData(sTemplatePath, "CWCustomProperty.UpdateShortDescription",oCaseWareApp);
			var sType = "U";
			var iDisplay = 1;
			var sKey = "updateClientFile_TR";
			var sTemplateDescrLong = CQSGetMetaData(sTemplatePath, "CWCustomProperty.UpdateLongDescription",oCaseWareApp);
			//Nico von Ronge - 23_06_2011 - Needed longer extended descriptions
			var sAddUpdateDescr = CQSGetMetaData(sTemplatePath, "CWCustomProperty.UpdateLongDescription1",oCaseWareApp);
		   // sTemplateDescrLong += sAddUpdateDescr.trim();
			sAddUpdateDescr = CQSGetMetaData(sTemplatePath, "CWCustomProperty.UpdateLongDescription2",oCaseWareApp);
		   // sTemplateDescrLong += sAddUpdateDescr.trim();
			sAddUpdateDescr = CQSGetMetaData(sTemplatePath, "CWCustomProperty.UpdateLongDescription3",oCaseWareApp);
		   // sTemplateDescrLong += sAddUpdateDescr.trim();
			var sTemplateDescrShortExtended = CQSGetMetaData(sTemplatePath, "CWCustomProperty.UpdateShortDescriptionExtended",oCaseWareApp);

			if(sTemplateProduct=="IFRS for SME") sTemplateProduct = "SME";
			if(sTemplateProduct=="IFRS") sTemplateProduct = "FULL_IFRS";
			if(sTemplateProduct=="IFRS7") sTemplateProduct = "FULL_IFRS";
			//now that we have all the details, lets switch off the other items in the
			//array and add the new item

			var sUpdate = ""
			var sFolder = ""
			//var sErrorMessage = willUpdatesRun(sUpdate, sFolder)
			sErrorMessage = ""
			var l_array = new Array(sKey,sTemplateVersion,sTemplateDescr,sType, iDisplay, sTemplateDescrLong, sClientVersion, sErrorMessage, sTemplateProduct,"Blank",sTemplateDescrShortExtended);
			aCurrentUpdates.push(l_array);
		}

		var bProbeUpdateAvailable = PROBEIsanUpdateAvailable(sClientProduct, sClientVersion, sClientLanguage, sCountry,oCaseWareApp,sClientFile_Path);
		
		if(bProbeUpdateAvailable)
		{
			//sCode, sProduct, sVersion, bAlternate, sLanguage, sCountryParam,sPathandFileName,sClientSubProduct,oCaseWareApp,sClientFile_Path
			var sTemplatePath = GetTemplatePath("", sClientProduct, sClientVersion, 0, sClientLanguage, sCountry,sPathandFileName,sClientSubProduct,oCaseWareApp,sClientFile_Path);
			var sTemplateProduct =  CQSGetMetaData(sTemplatePath, "CWCustomProperty.Probe_Product",oCaseWareApp);
			var sTemplateVersion = CQSGetMetaData(sTemplatePath, "CWCustomProperty.Probe_Cversion",oCaseWareApp);
			//create probe meta data for Short Description
			var sTemplateDescr = CQSGetMetaData(sTemplatePath, "CWCustomProperty.Probe_UpdateShortDescription",oCaseWareApp);
			var sType = "PU";
			var iDisplay = 1;
			var sKey = ""
			//this meta data item must still be added to the templates
			var sKey = CQSGetMetaData(sTemplatePath, "CWCustomProperty.Probe_ScriptFunction",oCaseWareApp);
			var sKey = "updateProbeClientFile_TR";
			//sKey = sKey + "_TR" //use this line of code if we get the function name from meta data
			var sTemplateDescrLong = CQSGetMetaData(sTemplatePath, "CWCustomProperty.Probe_UpdateLongDescription",oCaseWareApp);

			var sUpdate = ""
			var sFolder = "Probe"
		   // var sErrorMessage = willUpdatesRun(sUpdate, sFolder)
			var sErrorMessage = ""
			var l_array = new Array(sKey,sTemplateVersion,sTemplateDescr,sType, iDisplay, sTemplateDescrLong, sClientVersion, sErrorMessage, sTemplateProduct, sProbeClientVersion);
			aCurrentUpdates.push(l_array);
		}

		////debugger
			var bTaxConvertAvailable = TAXIsaConvertAvailable("",oCaseWareApp,sClientFile_Path);
			if(bTaxConvertAvailable)
			{
				var sTemplatePath = GetTemplatePath("", sClientProduct, sClientVersion, 0, sClientLanguage, sCountry,sPathandFileName,sClientSubProduct,oCaseWareApp,sClientFile_Path);
				var sTemplateProduct =  CQSGetMetaData(sTemplatePath, "CWCustomProperty.TC_Product",oCaseWareApp);

				var sTemplateVersion = CQSGetMetaData(sTemplatePath, "CWCustomProperty.TC_Version",oCaseWareApp);
				var sTemplateDescr = CQSGetMetaData(sTemplatePath, "CWCustomProperty.TC_UpdateShortDescription",oCaseWareApp);
				var sType = "TC";
				var iDisplay = 1;
				var sKey = "convertTaxClientFile_TR"
				var sTemplateDescrLong = CQSGetMetaData(sTemplatePath, "CWCustomProperty.TC_UpdateLongDescription",oCaseWareApp);

				var sUpdate = ""
				var sFolder = "TaxComp"
				//var sErrorMessage = willUpdatesRun(sUpdate, sFolder)
				var sErrorMessage = ""
				var l_array = new Array(sKey,sTemplateVersion,sTemplateDescr,sType, iDisplay, sTemplateDescrLong, sClientVersion, sErrorMessage, sTemplateProduct)
				aCurrentUpdates.push(l_array)
			}

		//debugger //GTS04012017
		aAvailableUpdates = aCurrentUpdates;
	   return aAvailableUpdates;
	}
	catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}finally{
		oClientFile = null;
		oSrcSystemCaseViewData = null;
	}
}

function TAXIsaConvertAvailable(bTemplateUpdate,oCaseWareApp,sClientFile_Path)
{
////debugger
//KN - Q2 2011
    try
    {
		if(!oCaseWareApp)
		{
			var oCaseWareApp = new ActiveXObject("CaseWare.Application");		
			var sClientFile_Path =  oCWApp.ActiveClient.FileName;	
        }
		//check if the client file has probe mmx meta data. only convert files that dont have probe mmx meta data
       // var sClientFilePath = sClientFile_Path;
        var sClientTaxProduct = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.TC_Product");
        //check if this machine has the probe mmx licence key in the registry. client must have a probe mmx licence for the conversion to take place
        var sTaxlicenceExists = true;//TAXdoesLicenceExist()

        //check if the client files corresponding template contains probe meta data
        //get the template file path
        var sCode = ""
        var sProduct = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Product");
        var sVersion = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Mapping");
        var sLanguage = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.ProductLanguage");
        var sCountry = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Country");
		var sPathandFileName = "";
		var sClientSubProduct = "";
        var sCQSMasterFilePath = GetTemplatePath(sCode, sProduct, sVersion, 0, sLanguage, sCountry,sPathandFileName,sClientSubProduct,oCaseWareApp,sClientFile_Path)

        //for a template update, set the master file path to the patch master
        if(bTemplateUpdate)
        {
            var sTempUpdatePath = getTempKLupdatePath(sProduct, sLanguage);
            var sCQSPatchMasterFilePath = sTempUpdatePath + "CQSUpgrade\\CQSUpgrade.ac";
            sCQSMasterFilePath = sCQSPatchMasterFilePath;
        }

        var sTemplateTaxProduct =  CQSGetMetaData(sCQSMasterFilePath, "CWCustomProperty.TC_Product",oCaseWareApp);

        var sClientTaxVersion =  CQSGetMetaData(sClientFile_Path, "CWCustomProperty.TC_Version",oCaseWareApp);
        var sTemplateTaxVersion =  CQSGetMetaData(sCQSMasterFilePath, "CWCustomProperty.TC_Version",oCaseWareApp);
        if(sClientTaxVersion < sTemplateTaxVersion)
             return true
        else
             return false
    }
    catch(e)
    {
		WScript.Echo("Error: "+e.description);
	}
}

function isTemplateInstalled(sProduct,oCaseWareApp,sClientFile_Path)
{
    try
    {


        var sCode;
        var sProductNameInMetaData;
        var sVersion;
        var sLanguage;
		if(!oCaseWareApp)
		{
			var oCWApp = new ActiveXObject("CaseWare.Application");
			var sClientFile_Path =  oCWApp.ActiveClient.FileName;
        }
		//get the client file object
        //var sClientFilePath =sClientFile_Path

        if(sProduct=="IFRS for SME") sProduct = "SME";
        if(sProduct=="IFRS") sProduct = "FULL_IFRS";
        if(sProduct=="IFRS7") sProduct = "FULL_IFRS";

        sCode = ""
        sVersion = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Mapping",oCaseWareApp);
        sProductNameInMetaData = sProduct;

        if(sVersion == "")
        {
            sVersion = "0.0";
        }
        //get the language of the client file
        if (sLanguage = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.ProductLanguage",oCaseWareApp)==4) //Afrikaans
        {
            var sLanguage = "4";
        }
        else //English
        {
            var sLanguage = "2";
        }

        var sCountry = CQSGetMetaData(sClientFile_Path, "CWCustomProperty.Country",oCaseWareApp);
        var sCQSMasterFilePath = GetTemplatePath(sCode, sProductNameInMetaData, sVersion, 0, sLanguage, sCountry,"","",oCaseWareApp,sClientFile_Path);

        //if the CQS master file related to this update does not exist exit the function/update
        if(sCQSMasterFilePath == "")
        {
            return 0
        }
        else
        {
            return 1
        }

    }
    catch(e)
    {
        WScript.Echo("Error: "+e.description);
    }
}

function getTempKLupdatePath(sProduct, sLanguage)
{
	try
	{
		if (!oDoc)
			var oDoc = document;

		var sUpdatePath = "";
		if(sProduct=="IFRS for SME") sProduct = "SME";
		if(sProduct=="IFRS") sProduct = "FULL_IFRS";
		if(sProduct=="IFRS7") sProduct = "FULL_IFRS";

		//determine which path to use to get to the update components for the relevant product
		if (sProduct == "FULL_IFRS")//IFRS
		{
			//check Language
			if (sLanguage == 4)//Afrikaans
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\IFRS\\AFR\\Temp\\"
			}
			else//English
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\IFRS\\Temp\\"
			}

			var sCountry = CQSGetClientCountry(oDoc)
			if (sCountry == AFRICA)
				sUpdatePath = document.interpret("prgdir()")+"Library\\IFRS Africa\\Temp\\"
		}
		else if(sProduct == "SME")//SME
		{
			//check Language
			if (sLanguage == 4)//Afrikaans
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\SME\\AFR\\Temp\\"
			}
			else//English
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\SME\\Temp\\"
			}
		}
		if(sProduct == "GRAP")//GRAP
		{
			//check Language
			if (sLanguage == 4)//Afrikaans
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\GRAP\\AFR\\Temp\\"
			}
			else//English
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\GRAP\\Temp\\"
			}
		}
		if(sProduct == "AuditINT")//AI
		{
			//check Language
			if (sLanguage == 4)//Afrikaans
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\AuditINT\\AFR\\Temp\\"
			}
			else//English
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\AuditINT\\Temp\\"
			}
		}
		return sUpdatePath
	}
	catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}
}


function getKLupdatePath(sProduct, sLanguage)
{
	try
	{
		if (!oDoc)
			var oDoc = document;

		var sUpdatePath = "";

		if(sProduct=="IFRS for SME") sProduct = "SME";
		if(sProduct=="IFRS") sProduct = "FULL_IFRS";
		if(sProduct=="IFRS7") sProduct = "FULL_IFRS";

		//determine which path to use to get to the update components for the relevant product
		if (sProduct == "FULL_IFRS")//IFRS
		{
			//check Language
			if (sLanguage == 4)//Afrikaans
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\IFRS\\AFR\\";
			}
			else//English
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\IFRS\\";
			}

			var sCountry = CQSGetClientCountry(oDoc)
			if (sCountry == AFRICA)
				sUpdatePath = document.interpret("prgdir()")+"Library\\IFRS Africa\\";
		}
		else if(sProduct == "SME")//SME
		{
			//check Language
			if (sLanguage == 4)//Afrikaans
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\SME\\AFR\\";
			}
			else//English
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\SME\\";
			}
		}

		if(sProduct == "GRAP")//GRAP
		{
			//check Language
			if (sLanguage == 4)//Afrikaans
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\GRAP\\AFR\\";
			}
			else//English
			{
				sUpdatePath = document.interpret("prgdir()")+"Library\\GRAP\\";
			}
		}
		
		if(sProduct == "IPSAS")//IPSAS
		{	
			sUpdatePath = document.interpret("prgdir()")+"Library\\IPSAS\\";
		}
		
		return sUpdatePath;
	}
	catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}
}



function getMetaData(sMetaDataName,sPath,oCaseWareApp)
{
	try
	{
		if(!oCaseWareApp)
			var oCaseWareApp = new ActiveXObject("CaseWare.Application")
		
		var oMetaData = oCaseWareApp.Clients.GetMetaData(sPath);
		var sReturnValue = ""
		
		if(oMetaData)
		{
			var oEnumerator = new Enumerator(oMetaData)
			for (;!oEnumerator.atEnd();oEnumerator.moveNext()) 
			{
				oItem = oEnumerator.item()
				var sCurrentName = oItem.Name
				var sCurrentValue = oItem.value
				
				if(sCurrentName==sMetaDataName)
					sReturnValue = sCurrentValue	
			}
			oMetaData = null
		}
		
		return sReturnValue
	}
	catch(e)
	{
		WScript.Echo("Error: "+e.description);
		success = false;
		//checkIfWeMustAbort(success)
	}
}

function GetTemplatePath(sCode, sProduct, sVersion, bAlternate, sLanguage, sCountryParam,sPathandFileName,sClientSubProduct,oCaseWareApp,sClientFile_Path)
{
	//if the template code has been passed back then we can 
	//get the information based on the code
	//If not we will run in a loop and retrieve the 
	//ac or ac_ file with the correct meta data
	try
	{

		//debugger
		//var sClientFile = sPathandFileName
		//var sProduct = getMetaData("CWCustomProperty.Product",sClientFile)
		
		var sCountry = verifyCountry(sCountryParam)
		
		var sPathandFileName = "";
		var l_Version = "0.0";
		var l_BigVersion = sVersion;
		
		if(!oCaseWareApp)
		{	
			var oCaseWareApp = new ActiveXObject("CaseWare.Application");
		}
		var oTempInfos = oCaseWareApp.TemplateList;
		
		if(sLanguage=="") sLanguage = 2;
		
		for (var i = 1; i <= oTempInfos.Count; i++)
		{
			if((typeof(sCode)!="undefined") && (sCode!=""))
			{
			if(sCode==oTempInfos.Item(i).Id)
			{
				var ls_Path = oTempInfos.Item(i).FilePath;
				var ls_Name = oTempInfos.Item(i).Name;
				
				ls_Path = CQSGetFilePathLib(ls_Path);
				//ls_Path = ls_Path.substring(0, ls_Path.length - ls_Name.length);  
				if (CQSFolderExists(ls_Path))
				{
				sPathandFileName = oTempInfos.Item(i).FilePath + ".AC";
					if(IsFileCompressed(sPathandFileName)) 
					sPathandFileName += "_";                 
					if (!CQSFileExists(sPathandFileName))    
					sPathandFileName = "";

				}
				break;
			}
			}
			else
			{
			//get the path and file name, then compare the product and version
			var ls_Path = oTempInfos.Item(i).FilePath;
			var ls_Name = oTempInfos.Item(i).Name;   
			ls_Path = CQSGetFilePathLib(ls_Path);
			//ls_Path = ls_Path.substring(0, ls_Path.length - ls_Name.length);	      
			if (CQSFolderExists(ls_Path))
			{
				var l_sPathandFileName =  oTempInfos.Item(i).FilePath + ".AC";
				if(IsFileCompressed(l_sPathandFileName)) 
				l_sPathandFileName += "_";
				if (CQSFileExists(l_sPathandFileName))
				{
					//return l_sPathandFileName;
					if(bAlternate) 
				{
					var l_Version = getMetaData("CWCustomProperty.Alt_Version",l_sPathandFileName,oCaseWareApp);
					var l_Product = getMetaData("CWCustomProperty.Alt_Product",l_sPathandFileName,oCaseWareApp);
					var l_Language = getMetaData("CWCustomProperty.AltProductLanguage",l_sPathandFileName,oCaseWareApp);
					var l_Country = getMetaData("CWCustomProperty.AltCountry",l_sPathandFileName,oCaseWareApp);
					var l_Framework = getMetaData("CWCustomProperty.Framework",l_sPathandFileName,oCaseWareApp);
				}
				else
				{
					var l_Version = getMetaData("CWCustomProperty.Mapping",l_sPathandFileName,oCaseWareApp);
					var l_Product = getMetaData("CWCustomProperty.Product",l_sPathandFileName,oCaseWareApp);
					var l_Language = getMetaData("CWCustomProperty.ProductLanguage",l_sPathandFileName,oCaseWareApp)
					var l_Country = getMetaData("CWCustomProperty.Country",l_sPathandFileName,oCaseWareApp);
					var l_Framework = getMetaData("CWCustomProperty.Framework",l_sPathandFileName,oCaseWareApp);
				}
				if(l_Language=="") l_Language = 2;
				

					if((l_Product=="IFRS for SME")||(l_Product=="IFRS") || (l_Product=="IFRS7") )
					{
						if(l_Product=="IFRS for SME") l_Product = "SME";
						if(l_Product=="IFRS") l_Product = "FULL_IFRS";
						if(l_Product=="IFRS7") l_Product = "FULL_IFRS";
						if(bAlternate==false)
						{
							addMetaData("CWCustomProperty.Product", l_Product, l_sPathandFileName);
						}
						else
						{
							addMetaData("CWCustomProperty.Product", l_Product, l_sPathandFileName);
						}
					}

				//CM - Fix
				if(l_Country!=sCountry)
				{
						sCountry=l_Country;
				}
				
				
				if(l_Framework=="") l_Framework = "AFS"

				if((l_Version >= l_BigVersion)&&(l_Product==sProduct)&&(l_Language==sLanguage)&&(l_Country==sCountry)&&(l_Framework=="AFS"))
				{
					l_BigVersion = l_Version;
					sPathandFileName = l_sPathandFileName;
				}
				
				}
			}  
			}
		}
		//oCWApp = null; 
		return sPathandFileName;
	}
	catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}
}

function CQSFolderExists(sPathName)
{
//This will return a boolean stating if the CW file is compressed or not
//The path and file name with the .ac should be passed in. 
//if the ac file can be found false will return, otherwise true will return
 
	try
    {
      var bResult = false;
      var oFSO = new ActiveXObject("Scripting.FileSystemObject");
      if(oFSO)
      {
        if (oFSO.FolderExists(sPathName))
			bResult = true;
			oFSO = null;
      }
    }
    catch(e)
    {
      WScript.Echo("Error: "+e.description);
    }     
  return bResult;
}

function CQSGetFilePathLib(sPathandFileName)
{
//This will return the path of a specified file
//if c:\Program files\MyFile.txt was passed in
//it will return c:\Program files\
	try
    {
      var sResult = "";
      var oFSO = new ActiveXObject("Scripting.FileSystemObject");
      if(oFSO)
      {
        sResult = oFSO.GetParentFolderName(sPathandFileName);
		oFSO = null;
      }
    }
    catch(e)
    {
      WScript.Echo("Error: "+e.description);
    }     
  return sResult;
}
function IsFileCompressed(sPathandFileName)
{
//This will return a boolean stating if the CW file is compressed or not
//The path and file name with the .ac should be passed in. 
//if the ac file can be found false will return, otherwise true will return 

   try
    {
	  //first check if there is a _ in the file name, if there is remove it
	  if(sPathandFileName.substring(sPathandFileName.length -1,sPathandFileName.length) == "_")
	    sPathandFileName = sPathandFileName.substring(0, sPathandFileName.length -1);
		
      var bResult = false;
      var oFSO = new ActiveXObject("Scripting.FileSystemObject");
      if(oFSO)
      {
        if (oFSO.FileExists(sPathandFileName + "_"))
		{
			bResult = true;
		}
		oFSO=null;
      }
    }
    catch(e)
    {
      WScript.Echo("Error: "+e.description);
    }     
  return bResult;
}
function CQSFileExists(sPathandFileName)
{
//This will return a boolean stating if the CW file is compressed or not
//The path and file name with the .ac should be passed in. 
//if the ac file can be found false will return, otherwise true will return 
	try
    {
      var bResult = false;
      var oFSO = new ActiveXObject("Scripting.FileSystemObject");
      if(oFSO)
      {
        if (oFSO.FileExists(sPathandFileName))
			bResult = true;
			oFSO=null;
      }
    }catch(e)
    {
      WScript.Echo("Error: "+e.description);
    }     
  return bResult;
}

function CQSGetMetaData(sFilePath, sPropertyName,oCaseWareApp)
{
//This will retrieve meta information in the file and path 
//specified, the property name is the item that will be 
//returned
	//Check if logging or debugger has been turned on
	//Removing the call to avoid stack overflow
	//checkDebugLib();
  try
  {
    var oMetaValue = "";
	var iCaseWareAppLocal = 0;
	
	//if no file path return empty string
	if(!isInputValid(sFilePath))  
	  return oMetaValue;
  

	if(!oCaseWareApp){
		var oCaseWareApp = new ActiveXObject("CaseWare.Application");
		iCaseWareAppLocal = 1;
	}
	try
	{
		var oMetaData = oCaseWareApp.Clients.GetMetaData(sFilePath)
		if(oMetaData.Exists(sPropertyName))
			oMetaValue = oMetaData.item(sPropertyName).value;
		else{
			if(oMetaData.Exists("CWCustomProperty."+sPropertyName))
				oMetaValue = oMetaData.item("CWCustomProperty."+sPropertyName).value;
		}
		
		if(iCaseWareAppLocal==1)
			oCaseWareApp = null;				  
	}
	catch(e)
	{
		try
		{
			if(oMetaData.Exists(sPropertyName))
				oMetaValue = oMetaData.item("CWCustomProperty."+sPropertyName).value;
			
			if(iCaseWareAppLocal==1)
				oCaseWareApp = null;
			
		}
		catch(e)
		{
			if(iCaseWareAppLocal==1)
				oCaseWareApp = null;	
			
			WScript.Echo("Error: "+e.description);
		}
	}

  }
  catch(e)
  {
    WScript.Echo("Error: "+e.description);
  }
  return oMetaValue;
}
//debugger;
//debugger;

function readTextFile(sFileName)
{
	try{
		var iForReading = 1, sFileData, oFile;
		var oFileSystemObject = WScript.CreateObject("Scripting.FileSystemObject");
		if(oFileSystemObject)			
			oFile = oFileSystemObject.OpenTextFile(sFileName, iForReading);
			if(oFile)
				sFileData = oFile.ReadAll();
			
		return sFileData; 
	}catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}finally{
		oFileSystemObject = null;
		oFile = null;
	}
}

function checkIfFileExist(sFilePath)
{
	try{
		var bReturn = false;
		var oFileSystemObject = WScript.CreateObject("Scripting.FileSystemObject");
		if(oFileSystemObject && oFileSystemObject.FileExists(sFilePath))
		{
			bReturn = true;
		}
		return bReturn;
	}catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}
}

function checkIfFolderExists(sFolder)
{
	try{
		var bReturn = false;
		var oFileSystemObject = WScript.CreateObject("Scripting.FileSystemObject");
		if(oFileSystemObject && oFileSystemObject.FolderExists(sFolder))
		{
			bReturn = true;
		}
		return bReturn;
	}catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}	
}

function getFileExtension(sFilePath)
{
	try{
		var sFileExtension = "";
		var oFileSystemObject = WScript.CreateObject("Scripting.FileSystemObject");
		if(oFileSystemObject && oFileSystemObject.FileExists(sFilePath))
		{
			sFileExtension = oFileSystemObject.GetExtensionName(sFilePath);
		}
		return sFileExtension;
	}catch(e)
	{
		WScript.Echo("Error: "+e.description);
	}	
}
updateFile();