'VBA vLookup Functions where skuNum is located in the first column and range is correctly named Products

Function itemTitle(skuNum)
'retrieve Item title from SKU

    itemTitle = Application.WorksheetFunction.VLookup(skuNum, [Products], 2, 0)

End Function

Function itemDesc(skuNum)
'retrieve item description from SKU

    itemDesc = Application.WorksheetFunction.VLookup(skuNum, [Products], 4, 0)

End Function

Function itemUPC(skuNum)
'retrieve item upc from SKU, No character limit or formatting

    itemUPC = Application.WorksheetFunction.VLookup(skuNum, [Products], 3, 0)

End Function
