Sub CATMain()

'BrowseForFile
Const WINDOW_HANDLE = 0
Const NO_OPTIONS = 0

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.BrowseForFolder _
(WINDOW_HANDLE, "Select a folder:", NO_OPTIONS, "C:\")
Set objFolderItem = objFolder.Self
objPath = objFolderItem.Path


'Get the root of the CATProduct
Dim oRootProduct As Product
Set oRootProduct = CATIA.ActiveDocument.Product

'Recursive function localSaveAs
localSaveAs oRootProduct, objPath

End Sub

Function localSaveAs(oRootProductItem, objPath)
Dim subRootProduct As Product
For Each subRootProduct In oRootProductItem.Products
    toSave = subRootProduct.ReferenceProduct.Parent.Name
    CATIA.Documents.Item(toSave).SaveAs (objPath & "\" & i & toSave)
    localSaveAs subRootProduct, objPath
Next
End Function

