namespace LiXcelCore
open Microsoft.Office.Interop

type Api ()=
    member this.Hello (formula:Excel.Range) = formula.Formula
    member this.PrintFormula (cell:Excel.Range) =
         let tree = Parser.parse (cell.Formula.ToString())
         let sb = new System.Text.StringBuilder()
         (Parser.prettyPrint sb tree).ToString() |> box