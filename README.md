# Macro-Nombramiento-base
Codigo para nombrar una base de datos en excel


'NOMBRAMIENTO BASE DATOS PRINCPAL

Sub BD()


Dim Fili, Filu As Long
Dim Coli, Colu As Integer



V.Range("INI").CurrentRegion.Name = "BDPRINCIPAL"

Set BB = V.Range("INI")

Fili = BB.Row
Filu = BB.End(xlDown).Row
Coli = BB.Column
Colu = BB.End(xlToRight).Column

V.Range(V.Cells(Fili, Coli), V.Cells(Fili, Colu)).Name = "Encabezado"


V.Range(V.Cells(Fili + 1, Coli), V.Cells(Filu, Coli)).Name = "MES"
V.Range(V.Cells(Fili + 1, Coli + 1), V.Cells(Filu, Coli + 1)).Name = "CLIENTETITULAR"
V.Range(V.Cells(Fili + 1, Coli + 2), V.Cells(Filu, Coli + 2)).Name = "PRODUCTO"
V.Range(V.Cells(Fili + 1, Coli + 3), V.Cells(Filu, Coli + 3)).Name = "JOBNUMBER"
V.Range(V.Cells(Fili + 1, Coli + 4), V.Cells(Filu, Coli + 4)).Name = "TEMA"
V.Range(V.Cells(Fili + 1, Coli + 5), V.Cells(Filu, Coli + 5)).Name = "COLABORADOR"
V.Range(V.Cells(Fili + 1, Coli + 6), V.Cells(Filu, Coli + 6)).Name = "DEPARTAMENTO"
V.Range(V.Cells(Fili + 1, Coli + 7), V.Cells(Filu, Coli + 7)).Name = "HORAS"
V.Range(V.Cells(Fili + 1, Coli + 8), V.Cells(Filu, Coli + 8)).Name = "COSTO"
V.Range(V.Cells(Fili + 1, Coli + 9), V.Cells(Filu, Coli + 9)).Name = "COSTOTOTAL"
V.Range(V.Cells(Fili + 1, Coli + 10), V.Cells(Filu, Coli + 10)).Name = "TIPO"


End Sub

Sub BDPROVEEDORES()

Dim Fili, Filu As Long
Dim Coli, Colu As Integer


P.Range("INIP").CurrentRegion.Name = "BDPROVEEDORES"

Set BB = P.Range("INIP")

Fili = BB.Row
Filu = BB.End(xlDown).Row
Coli = BB.Column
Colu = BB.End(xlToRight).Column

P.Range(P.Cells(Fili, Coli), P.Cells(Fili, Colu)).Name = "EncabezadoP"


P.Range(P.Cells(Fili + 1, Coli), P.Cells(Filu, Coli)).Name = "CUENTA"
P.Range(P.Cells(Fili + 1, Coli + 1), P.Cells(Filu, Coli + 1)).Name = "ESTRATEGA"
P.Range(P.Cells(Fili + 1, Coli + 2), P.Cells(Filu, Coli + 2)).Name = "DISEÑADOR"
P.Range(P.Cells(Fili + 1, Coli + 3), P.Cells(Filu, Coli + 3)).Name = "CONTENT"
P.Range(P.Cells(Fili + 1, Coli + 4), P.Cells(Filu, Coli + 4)).Name = "CM"
P.Range(P.Cells(Fili + 1, Coli + 5), P.Cells(Filu, Coli + 5)).Name = "ANIMADOR"
P.Range(P.Cells(Fili + 1, Coli + 6), P.Cells(Filu, Coli + 6)).Name = "PRODUCTOR"
P.Range(P.Cells(Fili + 1, Coli + 7), P.Cells(Filu, Coli + 7)).Name = "EJECUTIVO"
P.Range(P.Cells(Fili + 1, Coli + 8), P.Cells(Filu, Coli + 8)).Name = "ANALYTICS"
P.Range(P.Cells(Fili + 1, Coli + 9), P.Cells(Filu, Coli + 9)).Name = "CMCONTENT"
P.Range(P.Cells(Fili + 1, Coli + 10), P.Cells(Filu, Coli + 10)).Name = "MESP"
P.Range(P.Cells(Fili + 1, Coli + 11), P.Cells(Filu, Coli + 11)).Name = "AÑOP"
P.Range(P.Cells(Fili + 1, Coli + 12), P.Cells(Filu, Coli + 12)).Name = "GASTOSP"
P.Range(P.Cells(Fili + 1, Coli + 13), P.Cells(Filu, Coli + 13)).Name = "VERTICEP"
P.Range(P.Cells(Fili + 1, Coli + 14), P.Cells(Filu, Coli + 14)).Name = "WHATAGRAPH"
P.Range(P.Cells(Fili + 1, Coli + 15), P.Cells(Filu, Coli + 15)).Name = "BANCO"
P.Range(P.Cells(Fili + 1, Coli + 16), P.Cells(Filu, Coli + 16)).Name = "PROJECT"

End Sub
