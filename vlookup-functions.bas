# VBA vLookup Functions where skuNum is located in the first column and range is correctly named Products

Function itemTitle(skuNum)
'retrieve Item title

    itemTitle = Application.WorksheetFunction.VLookup(skuNum, [Products], 2, 0)

End Function
