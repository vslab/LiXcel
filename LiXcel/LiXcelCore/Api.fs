namespace LiXcelCore
open Microsoft.Office.Interop
open Microsoft.FSharp.Quotations

module priv =
    open Microsoft.FSharp.Reflection
    open Microsoft.FSharp.Quotations.Patterns
    let resolveContext (ctx:Excel.Range) (addr:string) =
        ctx.Worksheet.Range(addr|>box)

    let rec listDependencies (ctx:Excel.Range) (expr:Expr) =
        let lista = ref []
        let rec addrec name (ctx:Excel.Range) (expr:Expr) =
            if not (List.exists (fun (n,_,_) -> name =n) !lista) then
                let deps = expr.GetFreeVars()
                deps |> Seq.iter (fun v ->
                    let cell = resolveContext ctx (v.Name)
                    let formula = cell.Formula |> unbox
                    let exp = Parser.parseExpr formula
                    addrec v cell exp )
                if (List.exists (fun (n,_,_) -> name =n) !lista) then
                    failwith "circular reference"
                else
                    lista := (name,ctx,expr)::!lista
        let deps = expr.GetFreeVars()
        deps |> Seq.iter (fun v ->
            let cell = resolveContext ctx (v.Name)
            let formula = cell.Formula |> unbox
            let exp = Parser.parseExpr formula
            addrec v cell exp )
        !lista
    let rec eval (env:Map<Var,float>) (expr:Expr) =
        match expr with
        | Value(v,t) -> v,t
        //| Coerce(e,t) -> eval ctx e
        //| NewObject(ci,args) -> ci.Invoke(evalAll ctx args)
        //| NewArray(t,args) -> 
        //    let array = Array.CreateInstance(t, args.Length) 
        //    args |> List.iteri (fun i arg -> array.SetValue(eval ctx arg, i))
        //    box array
        | NewUnionCase(case,args) -> FSharpValue.MakeUnion(case, evalAll env args),typeof<float list>
        //| NewRecord(t,args) -> FSharpValue.MakeRecord(t, evalAll ctx args)
        //| NewTuple(args) ->
        //    let t = FSharpType.MakeTupleType [|for arg in args -> arg.Type|]
        //    FSharpValue.MakeTuple(evalAll ctx args, t)
        //| FieldGet(Some(Value(v,_)),fi) -> fi.GetValue(v)
        //| PropertyGet(None, pi, args) -> pi.GetValue(null, evalAll ctx args)
        //| PropertyGet(Some(x),pi,args) -> pi.GetValue(eval ctx x, evalAll ctx args)
        | DerivedPatterns.SpecificCall <@ (~-) @> (a,b,c)->
            let a,_ = eval env c.Head
            (-(a |>unbox<float>))|>box,typeof<float>
        | DerivedPatterns.SpecificCall <@ (-) @> (a,b,c)->
            let a,_ = eval env c.Head
            let b,_ = eval env c.Tail.Head
            ((a |>unbox<float>)-(b |>unbox<float>))|>box,typeof<float>
        | DerivedPatterns.SpecificCall <@ (~+) @> (a,b,c)->
            eval env c.Head
        | DerivedPatterns.SpecificCall <@ (/) @> (a,b,c)->
            let a,_ = eval env c.Head
            let b,_ = eval env c.Tail.Head
            ((a |>unbox<float>)/(b |>unbox<float>))|>box,typeof<float>
        | Call(None,mi,args) ->
            let args = evalAll env args
            try 
                let rval = mi.Invoke(null,args )
                let rtype = mi.ReturnType
                rval,rtype
            with
                | :? System.Reflection.TargetInvocationException as ex -> raise ex.InnerException
        //| Call(Some(x),mi,args) -> mi.Invoke(eval ctx x, evalAll ctx args)
        | Var(v) -> env.Item(v)|>box,typeof<float>
        | arg -> raise <| System.NotSupportedException(arg.ToString())
    and evalAll env args = [|for arg in args -> eval env arg |> fst|]
    //and getVar ->
      
    let Eval (ctx:Excel.Range) (e:Expr) :unit->float =
        let cellList = listDependencies ctx e |> List.rev
        let rec evald (ctx:Map<Var,float>) = function
            | [] -> ctx
            | (var,ct,expr)::t ->
                evald (ctx.Add(var,eval ctx expr|> fst|> unbox)) t
        fun () ->
            let ctx = evald Map.empty cellList
            eval ctx e |> fst |> unbox
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
        let evaluation = priv.Eval cell expr
        evaluation().ToString() |>box
