module LiXcelCore.Parser
open Microsoft.FSharp.Quotations
open Microsoft.FSharp.Quotations.Patterns
//this code from http://www.developerfusion.com/article/123830/functional-cells-a-spreadsheet-in-f/

type token =
    | WhiteSpace
    | Symbol of char
    | OpToken of string
    | RefToken of int * int
    | StrToken of string
    | NumToken of float 

let (|Match|_|) pattern input =
    let m = System.Text.RegularExpressions.Regex.Match(input, pattern)
    if m.Success then Some m.Value else None

let toRef (s:string) =
    let col = int s.[0] - int 'A'
    let row = s.Substring 1 |> int
    col, row-1

let toToken = function
    | Match @"^\s+" s -> s, WhiteSpace
    | Match @"^\+|^\-|^\*|^\/"  s -> s, OpToken s
    | Match @"^=|^<>|^<=|^>=|^>|^<"  s -> s, OpToken s   
    | Match @"^\(|^\)|^\,|^\:|^%" s -> s, Symbol s.[0]   
    | Match @"^[A-Z]\d+" s -> s, s |> toRef |> RefToken
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

type arithmeticOp = Add | Sub | Mul | Div
type logicalOp = Eq | Lt | Gt | Le | Ge | Ne
type formula =
    | Neg of formula
    | ArithmeticOp of formula * arithmeticOp * formula
    | LogicalOp of formula * logicalOp * formula
    | Num of float
    | Ref of int * int
    | Range of int * int * int * int
    | Fun of string * formula list
    | LiteralString of string

(*
let rec (|Term|_|) = function
    | Sum(f1, (OpToken(LogicOp op))::Sum(f2,t)) -> Some(LogicalOp(f1,op,f2),t)
    | Sum(f1,t) -> Some (f1,t)
    | _ -> None
and (|LogicOp|_|) = function
    | "=" ->  Some Eq | "<>" -> Some Ne
    | "<" ->  Some Lt | ">"  -> Some Gt
    | "<=" -> Some Le | ">=" -> Some Ge
    | _ -> None
and (|Sum|_|) = function
    | RenameTerm(f1, t) ->      
        let rec aux f1 = function        
            | SumOp op::RenameTerm(f2, t) -> aux (ArithmeticOp(f1,op,f2)) t               
            | t -> Some(f1, t)      
        aux f1 t  
    | _ -> None
and (|SumOp|_|) = function 
    | OpToken "+" -> Some Add | OpToken "-" -> Some Sub 
    | _ -> None
and (|RenameTerm|_|) = function  
    | OpToken "-"::RenameTerm(f, t) -> Some(Neg f, t)
    | RenameFactor(f1, ProductOp op::RenameTerm(f2, t)) ->
        Some(ArithmeticOp(f1,op,f2), t)       
    | RenameFactor(f, t) -> Some(f, t)  
    | _ -> None    
and (|ProductOp|_|) = function
    | OpToken "*" -> Some Mul | OpToken "/" -> Some Div
    | _ -> None
and (|RenameFactor|_|) = function      
    | NumToken n::t -> Some(Num n, t)
    | RefToken(x1,y1)::(Symbol ':'::RefToken(x2,y2)::t) -> 
        Some(Range(min x1 x2,min y1 y2,max x1 x2,max y1 y2),t)  
    | RefToken(x,y)::t -> Some(Ref(x,y), t)
    | Symbol '('::Term(f, Symbol ')'::t) -> Some(f, t)
    | StrToken s::Tuple(ps, t) -> Some(Fun(s,ps),t)  
    | _ -> None
and (|Tuple|_|) = function
    | Symbol '('::Params(ps, Symbol ')'::t) -> Some(ps, t)  
    | _ -> None
and (|Params|_|) = function
    | Term(f1, t) ->
        let rec aux fs = function
            | Symbol ','::Term(f2, t) -> aux (fs@[f2]) t
            | t -> fs, t
        Some(aux [f1] t)
    | t -> Some ([],t)

let parse (s:string) =
    if s.Chars (0) = '=' then
        tokenize (s.Substring(1)) |> function 
        | Term(f,[]) -> f 
        | _ -> failwith "Failed to parse formula"
    else
        LiteralString s
*)
//[<ReflectedDefinition>]
type FunctionLibrary =
    static member GetCellValue (x:int) (y:int) =
        0.0
    static member SUM values =
        List.sum<float> values

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
    | RefToken (x1,y1)::t -> 
        let literalx = <@ x1@>
        let literaly = <@ y1@>
        Some (<@FunctionLibrary.GetCellValue %literalx %literaly@>,t)
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
    | Expression(e,Symbol ',':: ArgList(a,t)) -> Some (e.Raw::a,t)//Some (<@ %e::%a @>,t)
    | Expression(e,t) -> Some (e.Raw::[],t) //Some (<@ %e::[] @>,t)
    | _ -> None
and (|FunCall|_|) = function
    | StrToken(name) :: Symbol '(' :: ArgList(a,Symbol ')' :: t) ->
        let methodInfo = typeof<FunctionLibrary>.GetMethod(name,List.map (fun (x:Expr) -> x.Type) a|> List.toArray) 
        let e =
            if methodInfo <> null then
                Expr.Call(methodInfo, a)
            else
                let methodInfo = typeof<FunctionLibrary>.GetMethod(name,[|typeof<float list>|]) 
                if methodInfo <> null then
                    let rec aux = function
                        | [] -> <@@ []:float list @@>
                        | h::t -> let r = aux t in <@@ %%h::%%r : float list@@>
                    Expr.Call(methodInfo,[aux a])
                else
                    failwithf "Method Unknown: %s" name
            |> Expr.Cast
        Some(e,t)
        //Some(<@ ( FunctionLibrary.Invoke name %a) @>,t)
    | _ -> None
let evaluate (valueAt:int * int -> string) formula =
    let rec eval = function
        | Neg f -> - (eval f) 
        | ArithmeticOp(f1,op,f2) -> arithmetic op (eval f1) (eval f2)
        | LogicalOp(f1,op,f2) -> if logic op (eval f1) (eval f2) then 0.0 else -1.0
        | Num d -> d
        | Ref(x,y) -> valueAt(x,y) |> float
        | Range _ -> invalidOp "Expected in function"
        | Fun("SUM",ps) -> ps |> evalAll |> List.sum
        | Fun("IF",[condition;f1;f2]) -> 
            if (eval condition)=0.0 then eval f1 else eval f2 
        | Fun(_,_) -> failwith "Unknown function"
        | LiteralString s -> 0.0
    and arithmetic = function
        | Add -> (+) | Sub -> (-) | Mul -> (*) | Div -> (/)
    and logic = function         
        | Eq -> (=)  | Ne -> (<>)
        | Lt -> (<)  | Gt -> (>)
        | Le -> (<=) | Ge -> (>=)
    and evalAll ps =
        ps |> List.collect (function            
            | Range(x1,y1,x2,y2) ->
                [for x=x1 to x2 do for y=y1 to y2 do yield valueAt(x,y) |> float]
            | x -> [eval x]            
        )
    eval formula

let parseExpr s =
    match tokenize s with
    | FullExpression e -> e
    | _ -> let n = nan in <@ n @>
let references formula =
    let rec traverse = function
        | Ref(x,y) -> [x,y]
        | Range(x1,y1,x2,y2) -> 
            [for x=x1 to x2 do for y=y1 to y2 do yield x,y]
        | Fun(_,ps) -> ps |> List.collect traverse
        | ArithmeticOp(f1,_,f2) | LogicalOp(f1,_,f2) -> 
            traverse f1 @ traverse f2
        | _ -> []
    traverse formula

//this is mine
let rec prettyPrint (sb:System.Text.StringBuilder) : formula -> System.Text.StringBuilder = function
    | Neg x ->
        let sb = sb.Append "Neg("
        let sb = prettyPrint sb x
        sb.Append ")"
    | ArithmeticOp(f1,op,f2) ->
        let sb = sb.Append "ArithmeticOp("
        let sb = prettyPrint sb f1
        let opString =
            match op with
            | Add -> "+"
            | Div -> "/"
            | Mul -> "*"
            | Sub -> "-"
        let sb = sb.AppendFormat(",{0},",opString)
        let sb = prettyPrint sb f2
        sb.Append ")"
    | Fun(f,argarray) ->
        let sb = sb.Append "F"
        let sb = sb.Append f
        let sb = sb.Append "("
        let sb = argarray |> List.fold (fun sb e -> let sb = prettyPrint sb e in sb.Append "," ) sb
        sb.Append ")"
    | Ref (x,y) -> 
        let sb = sb.Append "CellRef("
        let sb = sb.Append x
        let sb = sb.Append ','
        let sb = sb.Append y
        sb.Append ")"
    | Num d -> 
        let sb = sb.Append "Num("
        let sb = sb.Append d
        sb.Append ")"
    | LiteralString s -> sb.AppendFormat("\"{0}\"", s)
    | _ -> sb.Append "Lol!"
