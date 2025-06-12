Attribute VB_Name = "TableManager"
Option Explicit

' ==========================================
' Module TableManager
' ------------------------------------------
' Ce module centralise la logique de manipulation des objets ListObject (tableaux Excel).
' Il gère la création, le nommage et la collecte des tableaux gérés par l'addin.
' ==========================================

Private Const MODULE_NAME As String = "TableManager"
Private Const TABLE_PREFIX As String = "EE_"

' Génère un nom unique pour un nouveau tableau basé sur le nom de la catégorie.
' Le nom est préfixé par "EE_" et suffixé par un numéro si nécessaire.
Public Function GetUniqueTableName(ByVal CategoryName As String) As String
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "GetUniqueTableName"
    
    Dim baseName As String
    baseName = TABLE_PREFIX & Utilities.SanitizeTableName(CategoryName)
    
    ' Essayer d'abord sans numéro
    GetUniqueTableName = baseName
    If Not TableExists(ThisWorkbook, GetUniqueTableName) Then Exit Function
    
    ' Sinon, ajouter un numéro jusqu'à trouver un nom disponible
    Dim i As Long
    For i = 1 To 999
        GetUniqueTableName = baseName & "_" & Format(i, "000")
        If Not TableExists(ThisWorkbook, GetUniqueTableName) Then Exit Function
    Next i
    
    ' Si on arrive ici, c'est qu'on n'a pas trouvé de nom disponible
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Impossible de générer un nom unique pour le tableau '" & CategoryName & "'"
    GetUniqueTableName = ""
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "table_manager_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la génération du nom du tableau"
    GetUniqueTableName = ""
End Function

' Vérifie si un tableau avec le nom spécifié existe déjà dans le classeur.
Private Function TableExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Const PROC_NAME As String = "TableExists"
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo ErrorHandler ' Reset error handling after expected error
        If Not lo Is Nothing Then
            TableExists = True
            Exit Function
        End If
        Set lo = Nothing ' Réinitialiser pour la prochaine itération
    Next ws
    
    TableExists = False
    Exit Function
ErrorHandler:
    TableExists = False ' En cas d'erreur, considérer que la table n'existe pas
    ' It's debatable whether to log an error here if the goal is just to check existence.
    ' For now, let's assume a failure to check is an error condition.
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur dans TableExists pour le nom: " & tableName
End Function

' Collecte tous les tableaux gérés par l'addin dans le classeur spécifié.
' Un tableau est considéré comme "géré" s'il a le préfixe "EE_" et un commentaire de métadonnées.
Public Function CollectManagedTables(ByVal wb As Workbook) As Collection
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "CollectManagedTables"
    
    Set CollectManagedTables = New Collection
    
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim tableInfo As Object ' Scripting.Dictionary
    
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            ' Vérifier si c'est un tableau géré
            If Left(lo.Name, 3) = TABLE_PREFIX Then
                ' Vérifier le commentaire
                Dim hasComment As Boolean
                hasComment = False
                On Error Resume Next ' Expecting an error if comment doesn't exist
                hasComment = (Len(lo.Range.Cells(1, 1).Comment.Text) > 0)
                On Error GoTo ErrorHandler ' Reset error handling
                
                If hasComment Then
                    ' Créer un dictionnaire avec les infos du tableau
                    Set tableInfo = CreateObject("Scripting.Dictionary")
                    tableInfo.Add "Name", lo.Name
                    tableInfo.Add "SheetName", ws.Name
                    CollectManagedTables.Add tableInfo
                End If
            End If
        Next lo
    Next ws
    
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "table_manager_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la collecte des tableaux gérés"
    Set CollectManagedTables = New Collection
End Function

' Compte le nombre de tableaux gérés dans le classeur spécifié.
Public Function CountManagedTables(ByVal wb As Workbook) As Long
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "CountManagedTables"
    
    CountManagedTables = CollectManagedTables(wb).Count
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "table_manager_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur lors du comptage des tableaux gérés"
    CountManagedTables = 0
End Function