Attribute VB_Name = "TableMetadata"
Option Explicit

' ==========================================
' Module TableMetadata
' ------------------------------------------
' Ce module gère la sérialisation et désérialisation des métadonnées stockées dans les commentaires des tableaux.
' Il est utilisé pour stocker et récupérer les informations de chargement des tableaux.
' ==========================================

' --- Constantes du Module ---
Private Const MODULE_NAME As String = "TableMetadata"
Private Const META_DELIM As String = "|"
Private Const META_KEYVAL_DELIM As String = "="

' ==========================================
' Fonctions Publiques
' ==========================================

' Sérialise les informations de chargement en une chaîne de caractères pour le stockage.
Public Function SerializeLoadInfo(loadInfo As DataLoadInfo) As String
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "SerializeLoadInfo"
    
    Dim parts As Collection
    Set parts = New Collection
    
    parts.Add "CategoryName" & META_KEYVAL_DELIM & loadInfo.Category.CategoryName
    
    Dim sVals As String
    If Not loadInfo.SelectedValues Is Nothing Then
        If loadInfo.SelectedValues.Count > 0 Then
            Dim arrVals() As String
            ReDim arrVals(1 To loadInfo.SelectedValues.Count)
            Dim i As Long: i = 1
            Dim v As Variant
            For Each v In loadInfo.SelectedValues
                arrVals(i) = CStr(v)
                i = i + 1
            Next v
            sVals = Join(arrVals, ",")
        End If
    End If
    parts.Add "SelectedValues" & META_KEYVAL_DELIM & sVals
    
    parts.Add "ModeTransposed" & META_KEYVAL_DELIM & CStr(loadInfo.ModeTransposed)
    
    Dim tempArray() As String
    ReDim tempArray(1 To parts.Count)
    Dim j As Long
    For j = 1 To parts.Count
        tempArray(j) = parts(j)
    Next j

    SerializeLoadInfo = Join(tempArray, META_DELIM)
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur de sérialisation"
    SerializeLoadInfo = ""
End Function

' Désérialise une chaîne de caractères en un objet DataLoadInfo.
Public Function DeserializeLoadInfo(ByVal metadata As String) As DataLoadInfo
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "DeserializeLoadInfo"
    
    Dim outLoadInfo As DataLoadInfo
    
    ' S'assurer que les catégories sont initialisées avant de chercher dedans.
    If CategoryManager.CategoriesCount = 0 Then CategoryManager.InitCategories
    
    ' Initialiser la structure de sortie
    Set outLoadInfo.SelectedValues = New Collection
    outLoadInfo.ModeTransposed = False
    
    ' Découper la chaîne en parties
    Dim parts() As String
    parts = Split(metadata, META_DELIM)
    
    Dim part As Variant
    For Each part In parts
        Dim kvp() As String
        kvp = Split(part, META_KEYVAL_DELIM)
        
        If UBound(kvp) < 1 Then GoTo NextPart
        
        Select Case kvp(0)
            Case "CategoryName"
                ' Retrouver la catégorie complète à partir du nom
                outLoadInfo.Category = CategoryManager.GetCategoryByName(kvp(1))
                
            Case "SelectedValues"
                If kvp(1) <> "" Then
                    Dim vals() As String
                    vals = Split(kvp(1), ",")
                    Dim val As Variant
                    For Each val In vals
                        outLoadInfo.SelectedValues.Add val
                    Next val
                End If
                
            Case "ModeTransposed"
                outLoadInfo.ModeTransposed = (LCase(kvp(1)) = "true")
        End Select
NextPart:
    Next part
    
    Set DeserializeLoadInfo = outLoadInfo
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur de désérialisation"
    ' En cas d'erreur, retourner un objet vide pour éviter les plantages
    Dim emptyInfo As DataLoadInfo
    Set DeserializeLoadInfo = emptyInfo
End Function 