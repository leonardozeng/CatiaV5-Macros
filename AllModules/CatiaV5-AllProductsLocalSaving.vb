'Copyright (c) 2015 Krzysztof Gorzynski <gorzynskikrzysztof@gmail.com>
'
'Permission to use, copy, modify, and distribute this software for any
'purpose with or without fee is hereby granted, provided that the above
'copyright notice and this permission notice appear in all copies.
'
'THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
'WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
'MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
'ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
'WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
'ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
'OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
'----------------------------------------------------------------------------
' Macro: CatiaV5-AllProductsLocalSaving.catvbs
' Version: 0.1
' Code: Catia VBS
' Purpose: 
' Autor: Krzysztof Górzyński
' Datum: 31/03/2015
'----------------------------------------------------------------------------
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
