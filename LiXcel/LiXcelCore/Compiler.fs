module LiXcelCore.Compiler
open Microsoft.Office.Interop
open Microsoft.FSharp.Quotations

type CompiledExpr =
    | Variable of Var
    | Value of float
    | Addition of CompiledExpr*CompiledExpr
    | Subtraction of  CompiledExpr*CompiledExpr
    | Multiplication of  CompiledExpr*CompiledExpr
    | Division of CompiledExpr*CompiledExpr
    | UnaryMinus of CompiledExpr
    | UnaryFunction of (float->float)*CompiledExpr
    | BinaryFunction of (float->float->float)*CompiledExpr*CompiledExpr
    | GenericFun of System.Reflection.MethodInfo*(CompiledExpr list)
    | GenericInvoke of System.Reflection.MethodInfo*(CompiledExpr list)
    | Bind of Var* CompiledExpr *CompiledExpr

let FastEval (e:CompiledExpr):float =
    let rec receval (env:Map<Var,float>) = function
        | Variable var -> env.[var]
        | Value v -> v
        | Addition (a,b) ->
            let a = receval env a
            let b = receval env b
            (a + b)
        | Subtraction (a,b) ->
            let a = receval env a
            let b = receval env b
            (a - b)
        | Multiplication (a,b) ->
            let a = receval env a
            let b = receval env b
            (a * b)
        | Division (a,b) ->
            let a = receval env a
            let b = receval env b
            (a / b)
        | UnaryMinus a ->
            let a = receval env a
            -a
        | UnaryFunction (f,a) ->
            let a = receval env a
            f a
        | BinaryFunction (f,a,b) ->
            let a = receval env a
            let b = receval env b
            f a b
        | GenericFun (mi,args) ->
            mi.Invoke(null,args|> Seq.map (box<<receval env) |> Seq.toArray)  |>unbox
        | GenericInvoke (mi,args) ->
            mi.Invoke(null,[|args|> List.map (receval env) |])  |>unbox
        | Bind (var,a,b) ->
            let a = receval env a
            receval (env.Add(var,a)) b
    receval Map.empty e

//transforms a list expression in an expression list.
//Undefined behavior if called with expressions which aren't lists
let rec listExprTransformer : Expr->Expr list = function
    | Patterns.NewUnionCase (_,args) ->
        match args with
        | [] -> []
        | h::t::[] -> h :: listExprTransformer t
        | arg -> raise <| System.NotSupportedException(arg.ToString())
    | arg -> raise <| System.NotSupportedException(arg.ToString())
let rec exprCompiler : Expr->CompiledExpr= function
    | Patterns.Value(v,t) -> v|> unbox |> Value
    | Patterns.Var(v) -> v |> Variable
    //| NewUnionCase(case,args) -> FSharpValue.MakeUnion(case, evalAll env args),typeof<float list>
    | DerivedPatterns.SpecificCall <@ (~+) @> (a,b,c)->
        exprCompiler c.Head
    | DerivedPatterns.SpecificCall <@ (~-) @> (a,b,c)->
        UnaryMinus (exprCompiler c.Head)
    | DerivedPatterns.SpecificCall <@ (+) @> (a,b,c)->
        Addition(exprCompiler c.Head,exprCompiler c.Tail.Head)
    | DerivedPatterns.SpecificCall <@ (-) @> (a,b,c)->
        Subtraction(exprCompiler c.Head,exprCompiler c.Tail.Head)
    | DerivedPatterns.SpecificCall <@ (*) @> (a,b,c)->
        Multiplication(exprCompiler c.Head,exprCompiler c.Tail.Head)
    | DerivedPatterns.SpecificCall <@ (/) @> (a,b,c)->
        Division(exprCompiler c.Head,exprCompiler c.Tail.Head)
    | Patterns.Call(None,mi,args) ->
        match args with
        | [] -> GenericFun(mi,List.map exprCompiler args)
        | h::[] when  h.Type = typeof<float list> ->
            GenericInvoke(mi,List.map exprCompiler (listExprTransformer args.Head))
        |_ -> GenericFun(mi,List.map exprCompiler args)
    | Patterns.Let (var,a,b) ->
        Bind(var,exprCompiler a,exprCompiler b)
    | arg -> raise <| System.NotSupportedException(arg.ToString())

module priv =
    open Microsoft.FSharp.Reflection
    open Microsoft.FSharp.Quotations.Patterns
    let resolveContext (ctx:Excel.Range) (addr:string) =
        let s = addr.Split([|'/'|],2)
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
        | Var(v) -> env.Item(v)|>box,typeof<float>
        | Patterns.Let (var,a,b) ->
            let v = eval env a |> fst |> unbox
            eval (env.Add(var,v)) b
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
        | arg -> raise <| System.NotSupportedException(arg.ToString())
    and evalAll env args = [|for arg in args -> eval env arg |> fst|]
    //and getVar ->
      
let Compile (ctx:Excel.Range) (e:Expr) :unit->float =
    let cellList = priv.listDependencies ctx e |> List.rev
    let rec evald e = function
        | (var,ct,expr)::t ->
            Expr.Let( var, expr ,(evald e t))
        | [] ->
            e
    let fold = evald e cellList
    let cexpr = exprCompiler fold
    fun () -> FastEval cexpr

let CompileCell (cell:Excel.Range) =
    let expr = cell.Formula |> unbox |> Parser.parseExpr  cell.Worksheet.Name
    Compile cell expr