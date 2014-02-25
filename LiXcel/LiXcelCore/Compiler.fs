module LiXcelCore.Compiler
open Microsoft.Office.Interop
open Microsoft.FSharp.Quotations

module priv =
    open Microsoft.FSharp.Reflection
    open Microsoft.FSharp.Quotations.Patterns
    let resolveContext (ctx:Excel.Range) (addr:string) =
        let s = addr.Split('/')
        if s.Length = 1 then
            ctx.Worksheet.Range(addr|>box)
        else
            let wb = ctx.Worksheet.Parent |>unbox<Excel.Workbook> 
            let sheet = wb.Worksheets.Item(s.[1] |> box) |> unbox<Excel.Worksheet>
            sheet.Range(s.[0])
    let rec listDependencies (ctx:Excel.Range) (expr:Expr) =
        let lista = ref []
        let rec addrec (name:Var) (ctx:Excel.Range) (expr:Expr) =
            if not (List.exists (fun (n,_,_) -> name =n) !lista) then
                let deps = expr.GetFreeVars()
                deps |> Seq.iter (fun v ->
                    let cell = resolveContext ctx (v.Name)
                    let formula = cell.Formula |> unbox
                    let exp = Parser.parseExpr  (cell.Worksheet.Name) formula
                    addrec v cell exp )
                if (List.exists (fun (n,_,_) -> name =n) !lista) then
                    failwith "circular reference"
                else
                    lista := (name,ctx,expr)::!lista
        let deps = expr.GetFreeVars()
        deps |> Seq.iter (fun v ->
            let cell = resolveContext ctx (v.Name)
            let formula = cell.Formula |> unbox
            let exp = Parser.parseExpr  (cell.Worksheet.Name) formula
            addrec v cell exp)
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
      
let Compile (ctx:Excel.Range) (e:Expr) :unit->float =
    let cellList = priv.listDependencies ctx e |> List.rev
    let rec evald (ctx:Map<Var,float>) = function
        | [] -> ctx
        | (var,ct,expr)::t ->
            evald (ctx.Add(var,priv.eval ctx expr|> fst|> unbox)) t
    fun () ->
        let ctx = evald Map.empty cellList
        priv.eval ctx e |> fst |> unbox

let CompileCell (cell:Excel.Range) =
    let expr = cell.Formula |> unbox |> Parser.parseExpr  cell.Worksheet.Name
    Compile cell expr