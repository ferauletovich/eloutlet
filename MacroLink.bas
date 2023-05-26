Attribute VB_Name = "NewMacros"
Sub Links()
Attribute Links.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Ventanas Alum Blanco Vidrio Entero"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="VentanasVidEnt"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Ventanas Alum Blanco Vidrio Repartido"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="VentanasVidRep"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Ventanas Aluminio Natural"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="VentanasNat"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
            With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Ventanas Aluminio Con "
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="VentanasCelo"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Puertas Placa Madera y Chapa"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="PPlacaMYC"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Ventanas Balc"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Balcon"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Puertas Aluminio Acanalada y Tubular"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="PuertasALYTU"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Puertas Placa Marco Aluminio"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="PlacaMA"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Despenseros"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Despenseros"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Modular Capilla"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Capilla"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Barras & Desayunadores"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Barrasydesay"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Mesas"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Mesas"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Comodas"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Comodas"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Roperos"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Roperos"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Bibliotecas"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Bibliotecas"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Bajo y Alacena"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Bajoyala"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Camas"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Camas"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Colchones"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Colchones"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Rajas"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Rajas"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
    
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Puertas Chapa Inyectada"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="PChapaInyectada"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Puertas Chapa Otras"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="PChapaOtras"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
    
        With Selection.Find
     .Forward = True
     .ClearFormatting
     .MatchWholeWord = True
     .MatchCase = False
     .Wrap = wdFindContinue
     .Execute FindText:="Ventiluz"
    End With
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Ventiluz"
        .DefaultSorting = wdSortByLocation
        .ShowHidden = False
    End With
    
    
End Sub

