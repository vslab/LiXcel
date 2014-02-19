namespace LiXcelCore
open Microsoft.Office.Interop

type Api ()=
    member this.Hello (formula:Excel.Range) = formula.Formula
