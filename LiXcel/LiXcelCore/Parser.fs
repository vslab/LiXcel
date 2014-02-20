module LiXcelCore.Parser
open Microsoft.FSharp.Quotations
//open Microsoft.FSharp.Quotations.Patterns
//this code from http://www.developerfusion.com/article/123830/functional-cells-a-spreadsheet-in-f/

type token =
    | WhiteSpace
    | Symbol of char
    | OpToken of string
    | RefToken of string
    | StrToken of string
    | NumToken of float 

let (|Match|_|) pattern input =
    let m = System.Text.RegularExpressions.Regex.Match(input, pattern)
    if m.Success then Some m.Value else None

let toToken = function
    | Match @"^\s+" s -> s, WhiteSpace
    | Match @"^\+|^\-|^\*|^\/"  s -> s, OpToken s
    | Match @"^=|^<>|^<=|^>=|^>|^<"  s -> s, OpToken s   
    | Match @"^\(|^\)|^\,|^\:|^%" s -> s, Symbol s.[0]   
    | Match @"^[A-Z]\d+" s -> s, s |> RefToken
    | Match @"^[A-Za-z]+" s -> s, StrToken s
    | Match @"^\d+(\.\d+)?|\.\d+" s -> s, s |> float |> NumToken
    | _ -> invalidOp ""

let tokenize s =
    let rec tokenize' index (s:string) =
        if index = s.Length then [] 
        else
            let next = s.Substring index 
            let text, token = toToken next
            token :: tokenize' (index + text.Length) s
    tokenize' 0 s
    |> List.choose (function WhiteSpace -> None | t -> Some t)

let vars = new System.Collections.Generic.Dictionary<string,Var>()

let rec (|FullExpression|_|) = function
    | OpToken "="::Expression(e,[]) -> Some e
//    | OpToken "+"::Expression(e,[]) -> Some e //looks like it's not needed on modern excel
    | NumToken d::[] -> Some (<@ d @>)
    | OpToken "-"::NumToken d::[] -> Some (let d = -d in <@ d @>)
    | NumToken d::Symbol '%'::[] -> Some (let d = d/100.0 in <@ d @>)
    | OpToken "-"::NumToken d::Symbol '%'::[] -> Some (let d = -d/100.0 in <@ d @>)
    | _ -> None
and (|NoNegation|_|) = function
    | NumToken d::t -> Some(<@ d @>,t)
//    | RefToken (x1,y1)::Symbol ':'::RefToken (x2,y2)::t -> Some (<@@ GetRange(@@>,t)
    | RefToken (addr)::t ->
          let v =
            lock(vars) (fun () ->
                try
                    vars.Item(addr)
                with
                | :? System.Collections.Generic.KeyNotFoundException ->
                    let v = Var(addr,typeof<float>,false)
                    vars.Add(addr,v)
                    v)
          Some (Expr.Var(v)|> Expr.Cast,t)
    | Symbol '(' :: Expression(e,Symbol ')'::t) -> Some (<@ ( %e )@>,t)
    | FunCall(e,t) -> Some(e,t)
    | _ -> None
and (|NoPercent|_|) = function
    | OpToken "-"::NoPercent(e,t) -> Some (<@ (- %e) @>,t)
    | NoNegation(e,t) -> Some (e,t)
    | _ -> None
and (|NoExponentiation|_|) = function
    | NoPercent(e, t) -> 
        let rec aux e = function
            | Symbol '%'::t -> let e,t = aux e t in  <@(%e/100.0)@>,t
            | t -> e,t
        Some (aux e t)
//    | NoPercent(e,Symbol '%':: t) -> Some (<@@(%%e/100)@@>,t)
//    | NoPercent(e,t) -> Some (e,t)
    | _ -> None
and (|NoMulDiv|_|) = function
    | NoExponentiation(e,t) ->
        let rec aux left = function
            | OpToken "^"::NoExponentiation(right,t) -> let newLeft = <@(%left ** %right)@> in aux newLeft t
            | t -> left,t
        Some (aux e t)
    | _ -> None
and (|NoAddSub|_|) = function
    | NoMulDiv(e,t) ->
        let rec aux left = function
            | OpToken "*"::NoMulDiv(right,t) -> let newLeft = <@(%left * %right)@> in aux newLeft t
            | OpToken "/"::NoMulDiv(right,t) -> let newLeft = <@(%left / %right)@> in aux newLeft t
            | t -> left,t
        Some (aux e t)
    | _ -> None
and (|NoConcat|_|) = function
    | NoAddSub(e,t) ->
        let rec aux left = function
            | OpToken "+"::NoAddSub(right,t) -> let newLeft = <@(%left + %right)@> in aux newLeft t
            | OpToken "-"::NoAddSub(right,t) -> let newLeft = <@(%left - %right)@> in aux newLeft t
            | t -> left,t
        Some (aux e t)
    | _ -> None
and (|NoLogic|_|) = function
    | NoConcat(e,t) -> Some(e,t) //concat not implemented yet
    | _ -> None
and (|Expression|_|) = function
    | NoLogic(e,t) ->
        let rec aux left = function
            | OpToken "="::NoLogic(right,t) -> let newLeft = <@(if %left = %right then 1.0 else 0.0)@> in aux newLeft t
            | OpToken "<"::NoLogic(right,t) -> let newLeft = <@(if %left < %right then 1.0 else 0.0)@> in aux newLeft t
            | OpToken ">"::NoLogic(right,t) -> let newLeft = <@(if %left > %right then 1.0 else 0.0)@> in aux newLeft t
            | OpToken "<="::NoLogic(right,t) -> let newLeft = <@(if %left <= %right then 1.0 else 0.0)@> in aux newLeft t
            | OpToken ">="::NoLogic(right,t) -> let newLeft = <@(if %left >= %right then 1.0 else 0.0)@> in aux newLeft t
            | OpToken "<>"::NoLogic(right,t) -> let newLeft = <@(if %left <> %right then 1.0 else 0.0)@> in aux newLeft t
            | t -> left,t
        Some (aux e t)
    | _ -> None
and (|ArgList|_|) = function
    | Expression(e,Symbol ',':: ArgList(al,ae,t)) -> Some (e.Raw::al,<@ %e::%ae @>,t)
    | Expression(e,t) -> Some (e.Raw::[],<@ [%e] @>,t)
    | _ -> None
and (|FunCall|_|) = function
    | StrToken(name) :: Symbol '(' :: ArgList(al,ae,Symbol ')' :: t) ->
        let methodInfo = typeof<FunctionLibrary>.GetMethod(name,List.map (fun (x:Expr) -> x.Type) al|> List.toArray) 
        let e =
            if methodInfo <> null then
                Expr.Call(methodInfo, al)
            else
                let methodInfo = typeof<FunctionLibrary>.GetMethod(name,[|typeof<float list>|])
                if methodInfo <> null then 
                    Expr.Call(methodInfo,[ae])
                else
                    failwithf "Method Unknown: %s" name
            |> Expr.Cast
        Some(e,t)
        //Some(<@ ( FunctionLibrary.Invoke name %a) @>,t)
    | _ -> None

let parseExpr s =
    match tokenize s with
    | FullExpression e -> e
    | _ -> let n = nan in <@ n @>