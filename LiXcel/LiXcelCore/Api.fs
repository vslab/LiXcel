namespace LiXcelCore
open Microsoft.Office.Interop

type Api ()=

    member this.Hello (formula:Excel.Range) = formula.Formula
    (*member this.PrintFormula (cell:Excel.Range) =
         let tree = Parser.parse (cell.Formula.ToString())
         let sb = new System.Text.StringBuilder()
         (Parser.prettyPrint sb tree).ToString() |> box
*)
    member this.GetExpr (cell:Excel.Range) =
        let expr = Parser.parseExpr (cell.Formula.ToString())
        expr.ToString() |> box
    member this.Tokenize (formula:Excel.Range) =
        let tokens = Parser.tokenize (formula.Formula.ToString().Substring(1))
        let sb = new System.Text.StringBuilder()
        List.iter (fun t -> sb.AppendLine (t.ToString()) |> ignore) tokens
        sb.ToString() |> box