namespace LiXcelCore
open Microsoft.Office.Interop
open Microsoft.FSharp.Quotations

module priv =
    open Microsoft.FSharp.Reflection
    open Microsoft.FSharp.Quotations.Patterns
    let resolveContext (ctx:Excel.Range) (addr:string) =
        ctx.Worksheet.Range(addr|>box)

    let rec listDependencies (name:Var) (ctx:Excel.Range) (expr:Expr) =
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
        addrec name ctx expr
        !lista
    let rec getDep (ctx:Excel.Range) (expr:Expr) =
        expr.GetFreeVars () |> Seq.map (fun x ->
            let nextctx = resolveContext ctx x.Name
            
            )
    let rec eval (ctx:Excel.Range) (env:Map<Var,float>) = function
        | Value(v,t) -> v
        //| Coerce(e,t) -> eval ctx e
        //| NewObject(ci,args) -> ci.Invoke(evalAll ctx args)
        //| NewArray(t,args) -> 
        //    let array = Array.CreateInstance(t, args.Length) 
        //    args |> List.iteri (fun i arg -> array.SetValue(eval ctx arg, i))
        //    box array
        | NewUnionCase(case,args) -> FSharpValue.MakeUnion(case, evalAll ctx args)
        //| NewRecord(t,args) -> FSharpValue.MakeRecord(t, evalAll ctx args)
        //| NewTuple(args) ->
        //    let t = FSharpType.MakeTupleType [|for arg in args -> arg.Type|]
        //    FSharpValue.MakeTuple(evalAll ctx args, t)
        //| FieldGet(Some(Value(v,_)),fi) -> fi.GetValue(v)
        //| PropertyGet(None, pi, args) -> pi.GetValue(null, evalAll ctx args)
        //| PropertyGet(Some(x),pi,args) -> pi.GetValue(eval ctx x, evalAll ctx args)
        | Call(None,mi,args) -> mi.Invoke(null, evalAll ctx args)
        //| Call(Some(x),mi,args) -> mi.Invoke(eval ctx x, evalAll ctx args)
        | Var(v) -> getVar
        | arg -> raise <| System.NotSupportedException(arg.ToString())
    and evalAll ctx args = [|for arg in args -> eval ctx var arg|]
    //and getVar ->
      

        //nodes. name,ctx,expr,expr.GetFreeVars
    let testina (radice:Excel.Range, =
    let mutable l = []
    let mutable s = [radice]
    while s <> [] do
        let n = s.Head
        s <- s.Tail
        l <- n :: l
        n

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