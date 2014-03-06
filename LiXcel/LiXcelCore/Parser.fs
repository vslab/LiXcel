module LiXcelCore.Parser
open Microsoft.FSharp.Quotations
//open Microsoft.FSharp.Quotations.Patterns
//this code from http://www.developerfusion.com/article/123830/functional-cells-a-spreadsheet-in-f/


(*
 Type system:
    - unknown
        - scalar
            - float
            - string
            - boolean
        - array (bidimensionale di scalar)

*)
type Scalar =
    | Number of float
    | Label of string
    | Boolean of bool
type Unknown =
    | Scalar of Scalar
    | Array of Scalar [,]

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
    | Match @"^=|^<>|^<=|^>=|^>|^<|^&"  s -> s, OpToken s   
    | Match @"^\(|^\)|^\,|^\:|^%|^\!" s -> s, Symbol s.[0]   
    //| Match @"^[A-Z]\d+" s -> s, s |> RefToken
    | Match @"^\$?[A-Z][A-Z]?\$?\d+" s -> s, s.Replace("$","") |> RefToken
    | RefFull (s,sheet,var) -> s, RefTokenFull (sheet,var.Replace("$",""))
    //| Match @"^(([A-Za-z_][A-Za-z0-9_\.]*|\'([^\']|\'\')*\')\!)[A-Z][A-Z]?\d+" s -> s, s |> RefTokenFull
    | Match @"^[A-Za-z_][_\.A-Za-z0-9]*" s -> s, StrToken s
    | Match @"^\d+(\.\d+)?|^\.\d+" s -> s, s |> float |> NumToken
    | Match @"^\'[^\']*\'" s -> s,StrToken s
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

let vars = new System.Collections.Generic.Dictionary<string*string,Var>()
let getVar sheet name =
    lock(vars) (fun () ->
        try
            vars.Item( sheet,name)
        with
        | :? System.Collections.Generic.KeyNotFoundException ->
            let v = Var(name + "/" + sheet,typeof<Scalar>,false)
            vars.Add((sheet,name),v)
            v)

let forceFloatExpr (e:Expr) =
        match e.Type with
        | t when t = typeof<float> -> Some e
        | t when t = typeof<bool> -> Some <@@ if (%%e:bool) then 1.0 else 0.0 @@>
        | t when t = typeof<Scalar [,]> ->
            Some <@@ match (%%e:Scalar [,]).[0,0] with
                        | Number n -> n
                        | Boolean b -> if b then 1.0 else 0.0
                        | _ -> 0.0 @@>
        | _ -> None
let forceMinor (left:Expr) (right:Expr) =
    match left.Type,right.Type with
    | t when t= (typeof<bool>,typeof<bool>) -> Some <@@ (%%left:bool) < (%%right:bool) @@>
    | t,_ when t= typeof<bool> -> Some <@@ false @@>
    | _,t when t= typeof<bool> -> Some <@@ true @@>
    | t when t= (typeof<string>,typeof<string>) -> Some <@@ (%%left:string) < (%%right:string) @@>
    | t,_ when t= typeof<string> -> Some <@@ false @@>
    | _,t when t= typeof<string> -> Some <@@ true @@>
    | t when t= (typeof<float>,typeof<float>) -> Some <@@ (%%left:float) < (%%right:float) @@>
    | _ -> None
let forceMajor (left:Expr) (right:Expr) =
    match left.Type,right.Type with
    | t when t= (typeof<bool>,typeof<bool>) -> Some <@@ (%%left:bool) > (%%right:bool) @@>
    | t,_ when t= typeof<bool> -> Some <@@ true @@>
    | _,t when t= typeof<bool> -> Some <@@ false @@>
    | t when t= (typeof<string>,typeof<string>) -> Some <@@ (%%left:string) > (%%right:string) @@>
    | t,_ when t= typeof<string> -> Some <@@ true @@>
    | _,t when t= typeof<string> -> Some <@@ false @@>
    | t when t= (typeof<float>,typeof<float>) -> Some <@@ (%%left:float) > (%%right:float) @@>
    | _ -> None
let forceEqual (left:Expr) (right:Expr) =
    match left.Type,right.Type with
    | t when t= (typeof<bool>,typeof<bool>) -> Some <@@ (%%left:bool) = (%%right:bool) @@>
    | t,_ when t= typeof<bool> -> Some <@@ false @@>
    | _,t when t= typeof<bool> -> Some <@@ false @@>
    | t when t= (typeof<string>,typeof<string>) -> Some <@@ (%%left:string) = (%%right:string) @@>
    | t,_ when t= typeof<string> -> Some <@@ false @@>
    | _,t when t= typeof<string> -> Some <@@ false @@>
    | t when t= (typeof<float>,typeof<float>) -> Some <@@ (%%left:float) = (%%right:float) @@>
    | _ -> None
let parseExpr sheetName s = 
    let rec (|FullExpression|_|) = function
        | OpToken "="::Expression(e,[]) -> Some e
    //    | OpToken "+"::Expression(e,[]) -> Some e //looks like it's not needed on modern excel
        | NumToken d::[] -> Some (<@@ d @@>)
        | OpToken "-"::NumToken d::[] -> Some (let d = -d in <@@ d @@>)
        | NumToken d::Symbol '%'::[] -> Some (let d = d/100.0 in <@@ d @@>)
        | OpToken "-"::NumToken d::Symbol '%'::[] -> Some (let d = -d/100.0 in <@@ d @@>)
        | [] -> Some <@@ 0.0 @@>
        | _ -> None
    and (|NoNegation|_|) = function
        | NumToken d::t -> Some(<@@ d @@>,t)
    //    | RefToken (x1,y1)::Symbol ':'::RefToken (x2,y2)::t -> Some (<@@ GetRange(@@>,t)
        | RefToken (addr)::t ->
              Some (Expr.Var(getVar sheetName addr),t)
        | RefTokenFull (sheetName,addr)::t ->
              Some (Expr.Var(getVar sheetName addr),t)
        | Symbol '(' :: Expression(e,Symbol ')'::t) -> Some (<@@ ( %%e )@@>,t)
        | FunCall(e,t) -> Some(e,t)
        | _ -> None
    and (|NoPercent|_|) = function
        | OpToken "-"::NoPercent(e:Expr,t) ->
            match forceFloatExpr e with
            | None -> None 
            | Some e -> Some (<@@ -(%%e:float) @@>,t)
        | NoNegation(e,t) -> Some (e,t)
        | _ -> None
    and (|NoExponentiation|_|) = function
        | NoPercent(e, t) -> 
            let rec aux (e:Expr) = function
                | Symbol '%'::t ->
                    match aux e t with
                    | None -> None
                    | Some (e:Expr,t) ->
                        match forceFloatExpr e with
                        | None -> None
                        | Some e -> Some (<@@ (%%e:float)/100.0 @@>,t)
                | t -> Some (e,t)
            aux e t
        | _ -> None
    and (|NoMulDiv|_|) = function
        | NoExponentiation(e,t) ->
            let rec aux left = function
                | OpToken "^"::NoExponentiation(right,t) -> 
                    match forceFloatExpr right with
                    | Some right -> let newLeft = <@@ (%%left:float) ** (%%right:float) @@> in aux newLeft t
                    | _ -> None
                | t -> Some(left,t)
            match forceFloatExpr e with
            | Some lefter -> aux lefter t
            | _ -> None
        | _ -> None
    and (|NoAddSub|_|) = function
        | NoMulDiv(e,t) ->
            let rec aux left = function
                | OpToken "*"::NoMulDiv(right,t) ->
                    match forceFloatExpr right with
                    | Some right -> let newLeft = <@@(%%left * %%right)@@> in aux newLeft t
                    | _ -> None
                | OpToken "/"::NoMulDiv(right,t) ->
                    match forceFloatExpr right with
                    | Some right -> let newLeft = <@@(%%left / %%right)@@> in aux newLeft t
                    | _ -> None
                | t -> Some(left,t)
            match forceFloatExpr e with
            | Some lefter -> aux lefter t
            | _ -> None
        | _ -> None
    and (|NoConcat|_|) = function
        | NoAddSub(e,t) ->
            let rec aux left = function
                | OpToken "+"::NoAddSub(right,t) ->
                    match forceFloatExpr right with
                    | Some right -> let newLeft = <@@(%%left + %%right)@@> in aux newLeft t
                    | _ -> None
                | OpToken "-"::NoAddSub(right,t) ->
                    match forceFloatExpr right with
                    | Some right -> let newLeft = <@@(%%left - %%right)@@> in aux newLeft t
                    | _ -> None
                | t -> Some(left,t)
            match forceFloatExpr e with
            | Some lefter -> aux lefter t
            | _ -> None
        | _ -> None
    and (|NoLogic|_|) = function
        | NoConcat(s1,OpToken "&"::NoLogic(s2:Expr,t)) when (s1.Type = typeof<string>) && (s2.Type = typeof<string>) ->
            Some(<@@ (%%s1:string) + (%%s2:string) @@>,t) 
        | NoConcat(e,t) -> Some(e,t) 
        | _ -> None
    and (|Expression|_|) = function
        | NoLogic(e,t) ->
            let rec aux left = function
                | OpToken "="::NoLogic(right,t) -> match forceEqual left right with None -> None | Some newLeft -> aux newLeft t
                | OpToken "<"::NoLogic(right,t) -> match forceMinor left right with None -> None | Some newLeft -> aux newLeft t
                | OpToken ">"::NoLogic(right,t) -> match forceMajor left right with None -> None | Some newLeft -> aux newLeft t
                | OpToken "<>"::NoLogic(right,t) -> match forceEqual left right with None -> None | Some newLeft -> (let newLeft = <@@ not (%%newLeft:bool) @@> in aux newLeft t)
                | OpToken ">="::NoLogic(right,t) -> match forceMinor left right with None -> None | Some newLeft -> (let newLeft = <@@ not (%%newLeft:bool) @@> in aux newLeft t)
                | OpToken "<="::NoLogic(right,t) -> match forceMajor left right with None -> None | Some newLeft -> (let newLeft = <@@ not (%%newLeft:bool) @@> in aux newLeft t)
                | t -> Some (left,t)
            aux e t
        | _ -> None
    and (|ArgList|_|) = function
        | Expression(e,Symbol ',':: ArgList(al,t)) -> Some (e::al,t)
        | Expression(e,t) -> Some (e::[],t)
        | _ -> None
    and (|FunCall|_|) = function
        | StrToken(name) :: Symbol '(' :: ArgList(al,Symbol ')' :: t) ->
            let types = List.map (fun (x:Expr) -> x.Type) al
            let methodInfo = typeof<FunctionLibrary>.GetMethod(name,types|> List.toArray) 
            let e =
                if methodInfo <> null then
                    Expr.Call(methodInfo, al)
                else
                    let isSubtype t1 t2 =
                        t1 = t2 || (t2 = typeof<Scalar[,]>) || ((t1 <> typeof<Scalar[,]>) && (t2 = typeof<Scalar>))
                    match types |> List.fold (fun (ltype:System.Type Option) (t:System.Type) ->
                                                match ltype,t with
                                                | None,t -> Some t
                                                | (Some t1),t2 when t1 = t2 -> Some t1 
                                                | (Some t),_ | (_,t) when t = typeof<Scalar [,]> -> Some typeof<Scalar [,]> 
                                                | _ -> Some typeof<Scalar> 
                                                ) None with
                    | None -> failwithf "Method Unknown: %s" name
                    | Some leastType ->
                    let listType = typeof<obj list>.GetGenericTypeDefinition().MakeGenericType([|leastType|])
                    let cons = listType.GetMethod("Cons")
                    let rec toListExp :Expr list -> Expr = function
                    | [] -> Expr.PropertyGet(listType.GetProperty("Empty"))
                    | h::t -> let t = toListExp t in Expr.Call(cons,[h;t])
                    let scalarCast (e:Expr) =
                        let numberCast (e:Expr) = Expr.Call(typeof<Scalar>.GetMethod("Number"),[e])
                        let boolCast (e:Expr) = Expr.Call(typeof<Scalar>.GetMethod("Boolean"),[e])
                        let stringCast (e:Expr) = Expr.Call(typeof<Scalar>.GetMethod("Label"),[e])
                        match e.Type with
                        | t2 when t2 = typeof<float> -> numberCast e
                        | t2 when t2 = typeof<bool> -> boolCast e
                        | t2 when t2 = typeof<string> -> stringCast e
                        | t2 when t2 = typeof<Scalar> -> e
                        | _ -> failwith "cast not allowed"
                    let arrayCast (e:Expr) = <@@ array2D [[(%%e:Scalar)]] @@>
                    let typeCast t (e:Expr) :Expr = 
                        match t,e.Type with
                        | t1,t2 when t1 = t2 -> e
                        | t1,_ when t1 = typeof<Scalar> -> scalarCast e
                        | t1,_ when t1 = typeof<Scalar [,]> -> scalarCast e |> arrayCast
                        | _ -> failwith "cast not allowed"

                    let types = seq {
                                yield leastType
                                if isSubtype leastType typeof<Scalar> then yield typeof<Scalar>
                                if isSubtype leastType typeof<Scalar[,]> then yield typeof<Scalar[,]>
                                } |> Seq.toList
                    let methods = (types |> Seq.map (fun x -> typeof<FunctionLibrary>.GetMethod(name,[|x|]))) |> Seq.zip types
                    match Seq.tryFind (((<>)null) <<snd) methods with 
                    | None -> failwithf "Method Unknown: %s" name
                    | Some (t,methodInfo) -> Expr.Call(methodInfo,[ al |> List.map (typeCast t) |>  toListExp ])
            Some(e,t)
            //Some(<@ ( FunctionLibrary.Invoke name %a) @>,t)
        | StrToken(name) ::  Symbol '(' :: Symbol ')' ::t ->
            let pm = typeof<FunctionLibrary>.GetProperty(name).GetGetMethod(false)
            let mm = typeof<FunctionLibrary>.GetMethod(name,[||])
            let mi = if pm <> null then pm else mm
            if mi <> null then Some(Expr.Call(mi,[]),t) else None
        | _ -> None
    let tokens = tokenize s
    match tokens with
    | FullExpression e -> e
    | _ -> failwith (sprintf "invalid expression: %O" tokens)