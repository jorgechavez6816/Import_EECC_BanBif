Sub Main
	IgnoreWarning(True)
	Call ReportReaderImport()	'D:\RUC1\DATA\Archivos fuente.ILB\2022_BBIF.pdf
	Call AppendField()	'K_BBIF2022.IMD
	Call ModifyField()		'K_BBIF2022.IMD
	Call Summarization()	'K_BBIF2022.IMD
	Client.CloseAll
	Dim pm As Object
	Dim SourcePath As String
	Dim DestinationPath As String
	Set SourcePath = Client.WorkingDirectory
	Set DestinationPath = "D:\RUC1\DATA\_EECC"
	Client.RunAtServer False
	Set pm = Client.ProjectManagement
	pm.MoveDatabase SourcePath + "K_BBIF2022.IMD", DestinationPath
	pm.MoveDatabase SourcePath + "K.1_Resumen_BBIF.IMD", DestinationPath
	Set pm = Nothing
	Client.RefreshFileExplorer
End Sub

' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport
	dbName = "K_BBIF2022.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\BBIF_CTA_CTE.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\2022_BBIF.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Anexar campo
Function AppendField
	Set db = Client.OpenDatabase("K_BBIF2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "PERIODO"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """20""+@RIGHT(PER;2)+@RIGHT(PER;2)"
	field.Length = 6
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modificar campo
Function ModifyField
	Set db = Client.OpenDatabase("K_BBIF2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "PERIODO"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """20""+@RIGHT(PER;2)+@LEFT(PER;2)"
	field.Length = 6
	task.ReplaceField "PERIODO", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Análisis: Resumen
Function Summarization
	Set db = Client.OpenDatabase("K_BBIF2022.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "PERIODO"
	task.AddFieldToSummarize "CUENTA"
	task.AddFieldToTotal "DEBITO"
	task.AddFieldToTotal "CREDITO"
	dbName = "K.1_Resumen_BBIF.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

