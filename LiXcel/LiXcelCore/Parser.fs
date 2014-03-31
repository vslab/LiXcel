module LiXcelCore.Parser
open Microsoft.FSharp.Quotations
//open Microsoft.FSharp.Quotations.Patterns
//this code from http://www.developerfusion.com/article/123830/functional-cells-a-spreadsheet-in-f/

type token =
    | WhiteSpace
    | Symbol of char
    | OpToken of string
    | RefToken of string
    | RefTokenFull of string * string
    | StrToken of string
    | NumToken of float 

let (|Match|_|) pattern input =
    let m = System.Text.RegularExpressions.Regex.Match(input, pattern)
    if m.Success then Some m.Value else None
let (|RefFull|_|) input =
    let m = System.Text.RegularExpressions.Regex.Match(input, @"^(?:(?:\$?([A-Za-z_][A-Za-z0-9_\.]*)|\'((?:[^\']|\'\')*)\')\!)\$?([A-Z][A-Z]?\$?\d+)" )
    if m.Success then
        //todo: dequote sheet string
        let s = m.Groups.[2].Value.Replace("''","'")
        Some (m.Value,(m.Groups.[1].Value+ s),m.Groups.[3].Value)
    else
        None

let toToken = function
    | Match @"^\s+" s -> s, WhiteSpace
    | Match @"^\+|^\-|^\*|^\/"  s -> s, OpToken s
    | Match @"^=|^<>|^<=|^>=|^>|^<"  s -> s, OpToken s   
    | Match @"^\(|^\)|^\,|^\:|^%|^\!" s -> s, Symbol s.[0]   
    //| Match @"^[A-Z]\d+" s -> s, s |> RefToken
    | Match @"^\$?[A-Z][A-Z]?\$?\d+" s -> s, s.Replace("$","") |> RefToken
    | RefFull (s,sheet,var) -> s, RefTokenFull (sheet,var.Replace("$",""))
    //| Match @"^(([A-Za-z_][A-Za-z0-9_\.]*|\'([^\']|\'\')*\')\!)[A-Z][A-Z]?\d+" s -> s, s |> RefTokenFull
    | Match @"^[A-Za-z_][_\.A-Za-z0-9]*" s -> s, StrToken s
    | Match @"^\d+(\.\d+)?|^\.\d+" s -> s, s |> float |> NumToken
    | Match @"^\'[^\']*\'" s -> s,StrToken s
    | t -> sprintf "Invalid token: %O" t |> invalidOp


let tokenize s =
    let rec tokenize' index (s:string) =
        if index = s.Length then [] 
        else
            let next = s.Substring index 
            let text, token = toToken next
            token :: tokenize' (index + text.Length) s
    tokenize' 0 s
    |> List.choose (function WhiteSpace -> None | t -> Some t)

let vars = new System.Collections.Generic.Dictionary<string*string,Var>()
let getVar sheet name =
    lock(vars) (fun () ->
        try
            vars.Item( sheet,name)
        with
        | :? System.Collections.Generic.KeyNotFoundException ->
            let v = Var(name + "/" + sheet,typeof<float>,false)
            vars.Add((sheet,name),v)
            v)
let parseExpr sheetName s = 
    let rec (|FullExpression|_|) = function
        | OpToken "="::Expression(e,[]) -> Some e
    //    | OpToken "+"::Expression(e,[]) -> Some e //looks like it's not needed on modern excel
        | NumToken d::[] -> Some (<@ d @>)
        | OpToken "-"::NumToken d::[] -> Some (let d = -d in <@ d @>)
        | NumToken d::Symbol '%'::[] -> Some (let d = d/100.0 in <@ d @>)
        | OpToken "-"::NumToken d::Symbol '%'::[] -> Some (let d = -d/100.0 in <@ d @>)
        | [] -> Some <@ 0.0 @>
        | _ -> None
    and (|NoNegation|_|) = function
        | NumToken d::t -> Some(<@ d @>,t)
        | RefToken (a1)::Symbol ':'::RefToken (a2)::t -> failwith "Range expression not supported:"
        | RefToken (addr)::t ->
              Some (Expr.Var(getVar sheetName addr)|> Expr.Cast,t)
        | RefTokenFull (sheetName,addr)::t ->
              Some (Expr.Var(getVar sheetName addr)|> Expr.Cast,t)
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
            //before that, check if it is an "if"
            if name = "IF" then
                try
                    let g::tru::fal::_ = al
                    Some(Expr.IfThenElse(Expr.Coerce(g,typeof<bool>),tru,fal)|>Expr.Cast,t)
                with
                | e -> None
            else
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
        | StrToken(name) ::  Symbol '(' :: Symbol ')' ::t ->
            let methodInfo = typeof<FunctionLibrary>.GetProperty(name).GetGetMethod(false)
            Some(Expr.Call(methodInfo,[])|>Expr.Cast,t)
        | _ -> None
    try
        let tokens = tokenize s
        match tokens with
        | FullExpression e -> e
        | _ -> failwith (sprintf "invalid expression: %O" (tokens.ToString()))
    with
    | e -> failwith (sprintf "%s%s%s" e.Message System.Environment.NewLine s) 