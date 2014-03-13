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
        let columnCount = destination.Columns.Count
        let bucketCount = destination.Rows.Count
        let simulator =  Compiler.CompileCell origin 
        let buckets = Array.init bucketCount (fun i -> i,0)
        let rangeMin = ref minRange
        let rangeMax = ref maxRange
        let range = ref nan
        let bucketWidth = ref nan
        let load () =
            let simulationResult = Array.init iterationCount (simulator<<ignore)
            if (System.Double.IsNaN !range) then // do this only the first time
                rangeMin := if System.Double.IsNaN minRange then Array.min simulationResult else minRange
                rangeMax := if System.Double.IsNaN maxRange then Array.max simulationResult else maxRange
                range := !rangeMax - !rangeMin
                bucketWidth := if !range = 0.0 then 1.0 else !range/(float bucketCount)
                let minmax =
                    seq {
                        let i = ref 0
                        while true do
                            let imin = (float (!i) ) * !bucketWidth + !rangeMin
                            let imax = imin + !bucketWidth
                            yield imin,imax
                            i := !i+1
                        }
                    |> Seq.take bucketCount
                    |> Seq.toList
                if columnCount > 2 then
                    let lmin,lmax = minmax |> List.unzip
                    ((destination.Columns.Item(columnCount-1|>box)):?>Excel.Range).set_Value(System.Type.Missing,(lmax |> Seq.map (fun x -> [box x] )|> array2D))
                    ((destination.Columns.Item(columnCount-2|>box)):?>Excel.Range).set_Value(System.Type.Missing,(lmin |> Seq.map (fun x -> [box x] )|> array2D))
                else if columnCount  = 2 then
                    let labels = minmax |> List.map (fun (imin,imax) -> [sprintf "%4f - %4f"  imin  imax|>box])
                    ((destination.Columns.Item(1|>box)):?>Excel.Range).set_Value(System.Type.Missing,labels|>array2D)

            let bucketsSimulation =
                simulationResult
                |> Seq.countBy (fun x ->
                    min ((x - !rangeMin) / !bucketWidth|>truncate |> int ) (bucketCount - 1) )|> Seq.filter (fun (i,v) -> i>=0 && i<bucketCount) |> Seq.sort |> Seq.toList
            bucketsSimulation |> List.iter (fun (i,v) -> let _, old = buckets.[i] in buckets.[i] <- i, old + v)
            //destination.Application
            try
                (destination.Columns.Item(columnCount |>box):?>Excel.Range).set_Value (System.Type.Missing,
                    let buckets = Array.sort buckets
                    seq {
                    let i = ref 0
                    for j,count in buckets do
                        while !i < j do
                            yield [0|>box]
                            i := !i + 1
                        yield [count]
                        i := !i + 1
                    } |> Seq.toList |> array2D)
            with
            | _ -> ()
        load

    member this.Simulate (origin:Excel.Range) (destination:Excel.Range) (iterationCount:int) (minRange:float) (maxRange:float) =
        this.SimulateThreaded (origin:Excel.Range) (destination:Excel.Range) (iterationCount:int) (minRange:float) (maxRange:float) ()