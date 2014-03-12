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
        let expr = Parser.parseExpr (cell.Formula.ToString()) (cell.Worksheet.Name)
        let evaluation = Compiler.Compile cell expr
        evaluation().ToString() |>box

    


    /// this member creates and returns a closure with the compilated excel expression that can be called many tiems 
    /// results will cumulate
    member this.SimulateThreaded (origin:Excel.Range) (destination:Excel.Range) (iterationCount:int) (minRange:float) (maxRange:float) =
        let bucketCount = destination.Rows.Count
        let simulator =  Compiler.CompileCell origin 
        let buckets = Array.init bucketCount (fun i -> i,0)
        let rangeMin = ref nan
        let rangeMax = ref nan
        let range = ref nan
        let bucketWidth = ref nan
        let load () =
            let simulationResult = Array.init iterationCount (simulator<<ignore)
            let printer (bucketNumber,count) =
                    if bucketNumber < 0 then () else
                    if bucketNumber + 1 > bucketCount then () else
                    let imin = (float bucketNumber ) * !bucketWidth + !rangeMin
                    let imax = imin + !bucketWidth
                    let label = sprintf "%4f - %4f"  imin  imax
                    let labelCell :Excel.Range = destination.Item(1 + bucketNumber|>box,1|>box)|>unbox
                    let countCell :Excel.Range = destination.Item(1 + bucketNumber|>box,2|>box)|>unbox
                    try
                        labelCell.Formula <- label |> box
                        countCell.Formula <- sprintf "%d" count |> box
                    with
                        | _ -> ()
            if (System.Double.IsNaN !range) then // do this only the first time
                rangeMin := if System.Double.IsNaN minRange then Array.min simulationResult else minRange
                rangeMax := if System.Double.IsNaN maxRange then Array.max simulationResult else maxRange
                range := !rangeMax - !rangeMin
                bucketWidth := if !range = 0.0 then 1.0 else !range/(float bucketCount)
                for  x in 0..(bucketCount-1) do printer (x, 0)
            let bucketsSimulation =
                simulationResult
                |> Seq.countBy (fun x ->
                    min ((x - !rangeMin) / !bucketWidth|>truncate |> int ) (bucketCount - 1) )|> Seq.filter (fun (i,v) -> i>=0 && i<bucketCount) |> Seq.sort |> Seq.toList
            bucketsSimulation |> List.iter (fun (i,v) -> let _, old = buckets.[i] in buckets.[i] <- i, old + v)
            //destination.Application
            buckets |> Seq.iter printer
        load

    member this.Simulate (origin:Excel.Range) (destination:Excel.Range) (iterationCount:int) (minRange:float) (maxRange:float) =
        this.SimulateThreaded (origin:Excel.Range) (destination:Excel.Range) (iterationCount:int) (minRange:float) (maxRange:float) ()