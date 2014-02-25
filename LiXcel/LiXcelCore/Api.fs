namespace LiXcelCore
open Microsoft.Office.Interop
open Microsoft.FSharp.Quotations

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
    member this.Eval (cell:Excel.Range) =
        let expr = Parser.parseExpr (cell.Formula.ToString())
        let evaluation = Compiler.Compile cell expr
        evaluation().ToString() |>box

    member this.Simulate (origin:Excel.Range) (destination:Excel.Range) (iterationCount:int) =
        let bucketCount = destination.Rows.Count
        let simulator =  Compiler.CompileCell origin 
        let simulationResult = Array.init iterationCount (simulator<<ignore)
        let rangeMin = Array.min simulationResult
        let rangeMax = Array.max simulationResult
        let range = rangeMax-rangeMin
        let bucketWidth = if range = 0.0 then 1.0 else range/(float bucketCount)
        let buckets =
            simulationResult
            |> Seq.countBy (fun x ->
                min ((x-rangeMin) / bucketWidth|>truncate |> int ) (bucketCount - 1) )|> Seq.sort |> Seq.toList
        let printer (bucketNumber,count) =
                let imin = (float bucketNumber )*bucketWidth + rangeMin
                let imax = imin + bucketWidth
                let label = sprintf "%4f - %4f"  imin  imax
                let labelCell :Excel.Range = destination.Item(1 + bucketNumber|>box,1|>box)|>unbox
                let countCell :Excel.Range = destination.Item(1 + bucketNumber|>box,2|>box)|>unbox
                labelCell.Formula <- label |> box
                countCell.Formula <- sprintf "%d" count |> box
        for  x in 0..(bucketCount-1) do printer (x, 0)
        buckets |> Seq.iter printer