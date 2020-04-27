Public Class GetConfigData


    Public Property reader As New System.Configuration.AppSettingsReader


    Public Property ServerURL = (reader.GetValue("ServerURL", GetType(String)))
    Public Property ServerURLProcessLog = (reader.GetValue("ServerURLProcessLog", GetType(String)))
    Public Property DocumentVault = (reader.GetValue("DocumentVault", GetType(String)))
    Public Property BAObject = (reader.GetValue("BAObject", GetType(String)))
    Public Property seeLogPopError = (reader.GetValue("seeLogPopError", GetType(String)))
    Public Property writingLogError = (reader.GetValue("writingLogError", GetType(String)))
    Public Property stateMergerWorkFlowError = (reader.GetValue("stateMergerWorkFlowError", GetType(String)))
    Public Property fileOpenError = (reader.GetValue("fileOpenError", GetType(String)))
    Public Property boardAgendaNotReadyError = (reader.GetValue("boardAgendaNotReadyError", GetType(String)))
    Public Property selectMatterError = (reader.GetValue("selectMatterError", GetType(String)))
    Public Property mergeSuccess = (reader.GetValue("mergeSuccess", GetType(String)))
    Public Property mergeCanceled = (reader.GetValue("mergeCanceled", GetType(String)))
    Public Property lessDocumentsAlert = (reader.GetValue("lessDocumentsAlert", GetType(String)))
    Public Property finishedProcess = (reader.GetValue("finishedProcess", GetType(String)))

    '----------Documents------------------------------------------------
    Public Property AgendaItemGeneral = (reader.GetValue("AgendaItemGeneral", GetType(String)))
    Public Property StaffArgumentGeneral = (reader.GetValue("StaffArgumentGeneral", GetType(String)))
    Public Property AgendaItemEff = (reader.GetValue("AgendaItemEff", GetType(String)))
    Public Property StaffArgumentEff = (reader.GetValue("StaffArgumentEff", GetType(String)))
    Public Property AgendaItemHaywood = (reader.GetValue("AgendaItemHaywood", GetType(String)))
    Public Property StaffArgumentHaywood = (reader.GetValue("StaffArgumentHaywood", GetType(String)))
    Public Property AgendaItemDR = (reader.GetValue("AgendaItemDR", GetType(String)))
    Public Property StaffArgumentDR = (reader.GetValue("StaffArgumentDR", GetType(String)))
    Public Property AgendaItemMembership = (reader.GetValue("AgendaItemMembership", GetType(String)))
    Public Property StaffArgumentMembership = (reader.GetValue("StaffArgumentMembership", GetType(String)))
    Public Property AgendaItemReeval = (reader.GetValue("AgendaItemReeval", GetType(String)))
    Public Property StaffArgumentReeval = (reader.GetValue("StaffArgumentReeval", GetType(String)))
    'CoverSheets
    Public Property CoverSheetA = (reader.GetValue("CoverSheetA", GetType(String)))
    Public Property CoverSheetB = (reader.GetValue("CoverSheetB", GetType(String)))
    Public Property CoverSheetC = (reader.GetValue("CoverSheetC", GetType(String)))

    'Placeholders
    Public Property Placeholder1 = (reader.GetValue("Placeholder1", GetType(String)))
    Public Property Placeholder2 = (reader.GetValue("Placeholder2", GetType(String)))


End Class
